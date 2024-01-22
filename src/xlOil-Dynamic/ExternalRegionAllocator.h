#pragma once
#include <xloil/WindowsSlim.h>
#include <map>
#include <set>
#include <cassert>

namespace xloil
{
  /// <summary>
  /// Allocator which returns memory regions in a specified address range.
  /// This is done by repeated calls to VirtualAlloc until one sticks and 
  /// the since allocator is also unoptimised so is not likely to have high
  /// performance. It's purpose is to provide space for dynamically written 
  /// thunks, which must have addresses in the range  
  /// [imageBase,  min(imageBase + MAXDWORD, address_max)] to be described 
  /// in the DLL export table.
  /// 
  /// The allocator keeps its data structures external to the allocated 
  /// memory. This allows for locking the page write permissions of the
  /// allocated memory using VirtualProtect which is good security practice
  /// since those regions need to be executable to act as thunks.
  /// </summary>
  class ExternalRegionAllocator
  {
  private:
    static constexpr unsigned MIN_BLOCKSIZE = 4;

    struct Block
    {
      short offset;
      bool operator<(const Block& that) const { return offset < that.offset; }
    };
    struct FreeBlock
    {
      char* chunk;
      Block block;
    };
    class Chunk
    {
    private:
      std::set<Block> _blocks;
      short _size;
      short _bytesUsed;

    public:
      Chunk(unsigned size) 
        : _size(short(size >> MIN_BLOCKSIZE)), _bytesUsed(0) 
      {}

      void* toAddress(void* chunkStart, Block block) const
      {
        return (char*)chunkStart + (size_t(block.offset) << MIN_BLOCKSIZE);
      }
      Block appendBlock(unsigned blockSize)
      {
        auto[i, success] = _blocks.emplace(Block{ _bytesUsed });
        _bytesUsed += short(blockSize >> MIN_BLOCKSIZE);
        return *i;
      }
      unsigned getBlockSize(typename decltype(_blocks)::iterator i) const
      {
        if (i == _blocks.end())
          return 0;
        auto offset = i->offset;
        if (++i == _blocks.end())
          return 0;
        return (i->offset - offset) << MIN_BLOCKSIZE;
      }
      auto findBlock(char* chunkStart, void* memPtr) const
      {
        auto block = Block{ (short)((char*)memPtr - chunkStart) >> MIN_BLOCKSIZE };
        return _blocks.find(block);
      }
      auto splitBlock(char* chunkStart, Block block, unsigned size)
      {
        auto[i, success] = _blocks.emplace(Block{ short(block.offset + (size >> MIN_BLOCKSIZE)) });
        auto blocksize = getBlockSize(i);
        return std::make_pair(blocksize, FreeBlock{ chunkStart, short(block.offset + (size >> MIN_BLOCKSIZE)) });
      }
      unsigned free(unsigned size)
      {
        return _bytesUsed -= short(size >> MIN_BLOCKSIZE);
      }
      unsigned available() const { return (_size - _bytesUsed) << MIN_BLOCKSIZE; }
      unsigned size() const { return _size << MIN_BLOCKSIZE; }
    };


  public:

    ExternalRegionAllocator(void* minAddress, void* maxAddress = (void*)UINTPTR_MAX)
      : _minAddress(minAddress)
    {
      SYSTEM_INFO si;
      GetSystemInfo(&si);
      _maxAddress = std::min(maxAddress, si.lpMaximumApplicationAddress);
      _pageSize = si.dwPageSize;
      assert(_maxAddress > _minAddress);
      // assert pagesize is a power of 2?
    }

    ~ExternalRegionAllocator()
    {
      for (auto i : _chunks)
        VirtualFree(i.first, 0, MEM_RELEASE);
    }

    auto alloc(unsigned bytesRequested)
    {
      bytesRequested = align(bytesRequested, 1 << MIN_BLOCKSIZE);
      auto iCurrent = _chunks.begin();
      if (iCurrent != _chunks.end() && iCurrent->second.available() >= bytesRequested)
      {
        auto newBlock = iCurrent->second.appendBlock(bytesRequested);
        return iCurrent->second.toAddress(iCurrent->first, newBlock);
      }

      auto iFree = _freeList.lower_bound(bytesRequested);
      if (iFree != _freeList.end())
      {
        auto[chunkData, chunk] = *_chunks.find(iFree->second.chunk);
        auto* foundAddress = chunk.toAddress(chunkData, iFree->second.block);
        if (iFree->first - bytesRequested > 1 << MIN_BLOCKSIZE)
          _freeList.emplace(
            chunk.splitBlock(chunkData, iFree->second.block, bytesRequested));
        else
          _freeList.erase(iFree);
        return foundAddress;
      }

      auto bytesToAlloc = align(bytesRequested, _pageSize);
      auto* allocated = _allocateChunk(bytesRequested);
      if (!allocated)
        throw _badAllocError;
      _chunks.emplace(std::make_pair((char*)allocated, Chunk(bytesToAlloc)));
      return alloc(bytesRequested);
    }

    auto free(void* memPtr)
    {
      auto iChunk = _chunks.lower_bound((char*)memPtr);
      if (iChunk == _chunks.end() || iChunk->first > memPtr) 
        --iChunk;

      auto [chunkData, chunk] = *iChunk;

      auto iBlock = chunk.findBlock(chunkData, memPtr);
      auto blockSize = chunk.getBlockSize(iBlock);

      if (chunk.free(blockSize) == 0)
      {
        // Remove all from free list
        auto chunkEnd = chunkData + chunk.size();
        auto iFree = _freeList.begin();
        while (iFree != _freeList.end())
        {
          if (iFree->second.chunk >= chunkData && iFree->second.chunk < chunkEnd)
            iFree = _freeList.erase(iFree);
          else
            ++iFree;
        }
        _chunks.erase(iChunk);
        VirtualFree(chunkData, 0, MEM_RELEASE);
      }
      else
        _freeList.emplace(std::make_pair(blockSize, FreeBlock{ chunkData, *iBlock }));
    }

  private:

    inline unsigned align(unsigned val, unsigned powerOf2)
    {
      const auto mask = powerOf2 - 1;
      return val + mask & ~(mask);
    }

    /// <summary>
    /// Searches for a page top down from maxAddress or the lowest chunk
    /// already allocated. Returns null if a page cannot be allocated
    /// above minAddress
    /// </summary>
    inline void* _allocateChunk(unsigned numBytes) const
    {
      void* p = nullptr;
      for (auto address = (char*)(!_chunks.empty() ? _chunks.begin()->first : _maxAddress);
        address > _minAddress && p == nullptr;
        address -= _pageSize)
          p = VirtualAlloc(address, numBytes, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
      return p;
    }

  private:
    // Map from allocated data to chunk descriptor
    std::map<char*, Chunk> _chunks; 
    // Map from block size to block locator
    std::multimap<unsigned, FreeBlock> _freeList; 
    void* _maxAddress;
    void* _minAddress;
    DWORD _pageSize;
    std::bad_alloc _badAllocError;
  };
}