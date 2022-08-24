/*
    BSD 3-Clause License
    
    Copyright (c) 2020, Tresorit Kft.
    All rights reserved.
    
    Redistribution and use in source and binary forms, with or without
    modification, are permitted provided that the following conditions are met:
    
    * Redistributions of source code must retain the above copyright notice, this
      list of conditions and the following disclaimer.
    
    * Redistributions in binary form must reproduce the above copyright notice,
      this list of conditions and the following disclaimer in the documentation
      and/or other materials provided with the distribution.
    
    * Neither the name of the copyright holder nor the names of its
      contributors may be used to endorse or promote products derived from
      this software without specific prior written permission.
    
    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
    AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
    IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
    DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
    FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
    DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
    SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
    CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
    OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
    OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

#include "rdc_fs_watcher.h"

#include <array>
#include <type_traits>
#include <iostream>
#include <string>

 // Information about a subscription for the changes of a single directory.
class WatchInfo final
{
private:
	std::unique_ptr<OVERLAPPED> overlapped;
	// Stores FILE_NOTIFY_INFORMATION, that is written by RDC() asynchronously.
	// Has to be destructed after the directory handle is closed!
	// RDC() requires a "pointer to the DWORD-aligned formatted buffer."
	std::aligned_storage_t<64 * 1024, sizeof(DWORD)> notifBuffer;
	static_assert(
		sizeof(WatchInfo::notifBuffer) <= 64 * 1024, "Must be smaller than RDC()'s network limit!");
	static_assert(sizeof(WatchInfo::notifBuffer) >= (sizeof(FILE_NOTIFY_INFORMATION) + (1 << 15)),
		"Must be able to store a long path.");

	std::wstring path;
	HandlePtr directory;
	bool watchTree;

	void processNotification(
		const FILE_NOTIFY_INFORMATION& notIf,
		std::set<std::pair<std::wstring, uint32_t>>& notifications) const;
	// The FileName might be in the short 8.3 form, so we try to get the long form.
	// Similar solution to libuv's.
	std::wstring tryToGetLongName(const std::wstring& pathName) const;

public:
	const int64_t rId;
	enum class State : uint8_t
	{
		Initialized, // No outstanding RDC() call.
		Listening, // RDC() call was made, and we're waiting for changes.
		PendingClose, // Directory handle was closed, and we're waiting for the "closing"
					  // notification on IOCP.
					  // Most of the time this is an "empty" notification, but sometimes it is a
					  // legitimate notification about a change. This is behavior is not documented
					  // explicitly.
	};
	std::atomic<State> state;

	WatchInfo(const int64_t& rId, std::unique_ptr<OVERLAPPED>&& overlapped,
		const std::wstring& path, HandlePtr&& dirHandle, bool watchTree = true);
	WatchInfo(const WatchInfo&) = delete;
	WatchInfo& operator=(const WatchInfo&) = delete;
	~WatchInfo();

	bool listen();
	void stop();
	bool canRun() const
	{
		return this->state != State::PendingClose;
	}

	std::set<std::pair<std::wstring, uint32_t>> processNotifications() const;
};

RdcFSWatcher::RdcFSWatcher(decltype(changeEvent) changeEvent, decltype(errorEvent) errorEvent)
	: stopped(false)
	, changeEvent(changeEvent)
	, errorEvent(errorEvent)
{
	this->iocp.reset(CreateIoCompletionPort(INVALID_HANDLE_VALUE, NULL, NULL, 1));
	if (!this->iocp) {
		throw std::runtime_error("Error when creating IOCP.");
	}
	this->loop = std::thread([this]() { this->eventLoop(); });
}

RdcFSWatcher::~RdcFSWatcher()
{
	this->stopEventLoop();
	this->loop.join();
}

void RdcFSWatcher::eventLoop()
{
	DWORD numOfBytes = 0;
	OVERLAPPED* ov = nullptr;
	ULONG_PTR compKey = 0;
	BOOL res = FALSE;
	while ((res = GetQueuedCompletionStatus(this->iocp.get(), &numOfBytes, &compKey, &ov, INFINITE))
		!= FALSE) {
		if (compKey != 0 && compKey == reinterpret_cast<ULONG_PTR>(this)) {
			// stop "magic packet" was sent, so we shut down:
			break;
		}
		else {
			this->processEvent(numOfBytes, ov);
		}
	}
	this->stopped.store(true);
	if (res != FALSE) {
		// IOCP is intact, so we clean up outstanding calls:
		std::lock_guard<std::mutex> lock(this->watchInfoMutex);
		const auto hasPending = [&lock, this]() {
			return std::find_if(this->watchInfos.cbegin(), this->watchInfos.cend(),
				[](decltype(watchInfos)::const_reference value) {
					return value.second.state == WatchInfo::State::PendingClose;
				})
				!= this->watchInfos.cend();
		};
		while (hasPending()
			&& (res = GetQueuedCompletionStatus(
				this->iocp.get(), &numOfBytes, &compKey, &ov, INFINITE))
			!= FALSE) {
			const auto watchInfoIt = this->watchInfos.find(ov);
			if (watchInfoIt != this->watchInfos.end()) {
				watchInfoIt->second.state = WatchInfo::State::PendingClose;
				this->watchInfos.erase(watchInfoIt);
			}
		}
	}
	else {
		// alert all subscribers that they will not receive events from now on:
		OutputDebugStringA(("RdcFSWatcher: error in event loop: " + std::to_string(GetLastError())).c_str());
		std::lock_guard<std::mutex> lock(this->watchInfoMutex);
		for (auto& watchInfo : this->watchInfos) {
			this->errorEvent(watchInfo.second.rId);
		}
	}
}

void RdcFSWatcher::stopEventLoop()
{
	{
		std::lock_guard<std::mutex> lock(this->watchInfoMutex);
		for (auto& watchInfo : this->watchInfos) {
			watchInfo.second.stop();
		}
	}
	if (this->iocp.get() != INVALID_HANDLE_VALUE) {
		// send stop "magic packet"
		PostQueuedCompletionStatus(this->iocp.get(), 0, reinterpret_cast<ULONG_PTR>(this), nullptr);
	}
}

void RdcFSWatcher::processEvent(DWORD numberOfBytesTrs, OVERLAPPED* overlapped)
{
	std::lock_guard<std::mutex> lock(this->watchInfoMutex);
	// initialization:
	const auto watchInfoIt = this->watchInfos.find(overlapped);
	if (watchInfoIt == this->watchInfos.end()) {
		OutputDebugStringA("WatchInfo was not found for filesystem event.");
		return;
	}
	if (watchInfoIt->second.state == WatchInfo::State::Listening) {
		watchInfoIt->second.state = WatchInfo::State::Initialized;
	}

	// actual logic:
	if (numberOfBytesTrs == 0) {
		if (watchInfoIt->second.state == WatchInfo::State::PendingClose) {
			// this is the "closing" notification, se we clean up:
			this->watchInfos.erase(watchInfoIt);
		}
		else {
			this->errorEvent(watchInfoIt->second.rId);
		}
		return;
	}

	WatchInfo& watchInfo = watchInfoIt->second;

	// If we're already in PendingClose state, and receive a legitimate notification, then
	// we don't emit a change notification, however, we delete the WatchInfo, just like when
	// we receive a "closing" notification.

	if (watchInfo.canRun()) {
		auto notificationResult = watchInfo.processNotifications();

		if (!notificationResult.empty()) {
			this->changeEvent(watchInfo.rId, std::move(notificationResult));
		}
		auto res = watchInfo.listen();
		if (!res) {
			this->errorEvent(watchInfo.rId);
			this->watchInfos.erase(watchInfoIt);
		}
	}
	else {
		this->watchInfos.erase(watchInfoIt);
	}
}

bool RdcFSWatcher::addDirectory(int64_t id, const std::wstring& path)
{
	if (this->stopped) {
		throw std::runtime_error("Watcher thread is not running.");
	}
	HandlePtr dirHandle(CreateFile(path.c_str(), FILE_LIST_DIRECTORY,
		FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING,
		FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED, NULL));
	if (dirHandle.get() == INVALID_HANDLE_VALUE) {
		throw std::runtime_error("Cannot create directory handle: " + std::to_string(GetLastError()));
		return false;
	}
	// check if it is even a directory:
	{
		BY_HANDLE_FILE_INFORMATION fileInfo{};
		const BOOL res = GetFileInformationByHandle(dirHandle.get(), &fileInfo);
		if (res == FALSE) {
			return false;
		}
		else if (!(fileInfo.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
			throw std::runtime_error("Not a directory.");
		}
	}
	// the "old" IOCP handle should not be freed, because:
	// https://devblogs.microsoft.com/oldnewthing/20130823-00/?p=3423
	if (!CreateIoCompletionPort(dirHandle.get(), this->iocp.get(), NULL, 1)) {
		throw std::runtime_error("Cannot create IOCP: " + std::to_string(GetLastError()));
	}

	// create the internal data structures, and set up listening with RDC():
	{
		auto ov = std::make_unique<OVERLAPPED>();
		OVERLAPPED* const ovPtr = ov.get();
		{
			std::lock_guard<std::mutex> lock(this->watchInfoMutex);
			const auto info = this->watchInfos.emplace(
				std::piecewise_construct, 
				std::forward_as_tuple(ovPtr),
				std::forward_as_tuple(id, std::move(ov), path, std::move(dirHandle)));
			if (!info.second) {
				return false;
			}
			if (this->stopped) {
				this->watchInfos.erase(info.first);
				return false;
			}
			else {
				bool result = info.first->second.listen();
				if (!result) {
					this->watchInfos.erase(info.first);
					return false;
				}
			}
		}
	}
	return true;
}

void RdcFSWatcher::removeDirectory(int64_t id)
{
	std::lock_guard<std::mutex> lock(this->watchInfoMutex);
	auto watchInfoIt = std::find_if(
		this->watchInfos.begin(), this->watchInfos.end(),
		[id](decltype(this->watchInfos)::const_reference value) {
			return value.second.rId == id;
		});
	if (watchInfoIt != this->watchInfos.end()) {
		watchInfoIt->second.stop();
	}
	// we should not emit events for this particular directory after this call
}

WatchInfo::WatchInfo(const int64_t& rId, std::unique_ptr<OVERLAPPED>&& overlapped,
	const std::wstring& path, HandlePtr&& dirHandle, bool watchTree)
	: overlapped(std::move(overlapped))
	, notifBuffer()
	, path(path)
	, directory(std::move(dirHandle))
	, rId(rId)
	, state(State::Initialized)
	, watchTree(watchTree)
{
}

WatchInfo::~WatchInfo()
{
	if (this->state == State::Listening) {
		OutputDebugStringA("Destructing a listening WatchInfo");
	}
}

bool WatchInfo::listen()
{
	if (this->state != State::Initialized) {
		throw std::runtime_error("WatchInfo: invalid state");
		return false;
	}
	constexpr DWORD flags = FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME
		| FILE_NOTIFY_CHANGE_SIZE | FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_CREATION
		| FILE_NOTIFY_CHANGE_SECURITY;
	const BOOL res = ReadDirectoryChangesW(this->directory.get(), &this->notifBuffer,
		static_cast<DWORD>(sizeof(this->notifBuffer)), watchTree, flags,
		nullptr /* lpBytesReturned */, this->overlapped.get(), nullptr /* lpCompletionRoutine */);
	if (res == FALSE) {
		this->state = State::Initialized;
		throw std::runtime_error("An error has occurred: " + std::to_string(GetLastError()));
	}
	this->state = State::Listening;
	return true;
}

std::wstring WatchInfo::tryToGetLongName(const std::wstring& pathName) const
{
	const std::wstring fullPath = this->path + L"\\" + pathName;
	const DWORD longSize = GetLongPathName(fullPath.c_str(), NULL, 0);
	std::wstring longPathName;
	longPathName.resize(longSize);
	const DWORD retVal = GetLongPathName(
		fullPath.c_str(), (wchar_t*)longPathName.data(), static_cast<DWORD>(longPathName.size()));
	if (retVal == 0) {
		return pathName;
	}
	while (!longPathName.empty() && longPathName.back() == L'\0') {
		longPathName.pop_back();
	}
	if (longPathName.find(this->path) == 0 && this->path.size() < longPathName.size()) {
		std::wstring longName = longPathName.substr(this->path.size() + 1);
		if (longName.empty()) {
			return pathName;
		}
		else {
			return longName;
		}
	}
	else {
		return pathName;
	}
}

void WatchInfo::stop()
{
	if (this->state == State::Listening) {
		this->state = State::PendingClose;
	}
	this->directory.reset(INVALID_HANDLE_VALUE);
}

std::set<std::pair<std::wstring, uint32_t>> WatchInfo::processNotifications() const
{
	std::set<std::pair<std::wstring, uint32_t>> notifications;

	auto notInf = reinterpret_cast<const FILE_NOTIFY_INFORMATION*>(&this->notifBuffer);
	for (bool moreNotif = true; moreNotif; moreNotif = notInf->NextEntryOffset > 0,
		notInf = reinterpret_cast<const FILE_NOTIFY_INFORMATION*>(
			reinterpret_cast<const char*>(notInf) + notInf->NextEntryOffset)) {
		this->processNotification(*notInf, notifications);
	}

	return std::move(notifications);
}

void WatchInfo::processNotification(
	const FILE_NOTIFY_INFORMATION& notIf,
	std::set<std::pair<std::wstring, uint32_t>>& notifications) const
{
	std::wstring wPathName(
		notIf.FileName, notIf.FileName + (notIf.FileNameLength / sizeof(notIf.FileName)));
	if (notIf.Action != FILE_ACTION_REMOVED && notIf.Action != FILE_ACTION_RENAMED_OLD_NAME) {
		std::wstring longName = this->tryToGetLongName(wPathName);
		if (longName != wPathName) {
			wPathName = longName;
		}
	}
	notifications.emplace(wPathName, notIf.Action);
}
