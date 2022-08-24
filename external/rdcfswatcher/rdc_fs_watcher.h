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

#pragma once

#include <atomic>
#include <thread>
#include <mutex>
#include <map>
#include <set>
#include <functional>

#include <windows.h>

class WatchInfo;

struct HandleDeleter
{
	void operator()(HANDLE handle)
	{
		if (handle != INVALID_HANDLE_VALUE && handle != NULL) {
			CloseHandle(handle);
		}
	}
};

using HandlePtr = std::unique_ptr<std::remove_pointer<HANDLE>::type, HandleDeleter>;

/**
 * Error handling is done by returning a bool, where true means success, and false means failure.
 * Additional error messages are written to std::cerr.
 */
class RdcFSWatcher final
{
  std::function<void(int64_t /* id */,
    const std::set<std::pair<std::wstring /* path */, uint32_t /* action */>>&)> changeEvent;
  std::function<void(int64_t)> errorEvent;

public:
	RdcFSWatcher(decltype(changeEvent) changeEvent, decltype(errorEvent) errorEvent);
	~RdcFSWatcher();

	void stopEventLoop();

	bool addDirectory(int64_t id, const std::wstring& path);
	void removeDirectory(int64_t id);

private:
	std::map<OVERLAPPED*, WatchInfo> watchInfos;
	std::mutex watchInfoMutex;

	// I/O completion port (IOCP)
	// It should be only "read" (using GetQueuedCompletionStatus()) from one thread.
	HandlePtr iocp;
	// has the "event loop" stopped?
	std::atomic<bool> stopped;
	// runs the "event loop", that checks the IOCP
	std::thread loop;

	void eventLoop();
	void processEvent(DWORD numberOfBytesTrs, OVERLAPPED* overlapped);
};
