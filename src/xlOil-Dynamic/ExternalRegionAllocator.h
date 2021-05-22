#pragma once
#include <xloil/WindowsSlim.h>
#include <list>
#include <map>
#include <vector>
#include <set>
#include <cassert>

namespace xloil
{
  /// <summary>
  /// Allocator which returns memory regions in a specified address range.
  /// This is done by repeated calls to VirtualAlloc until one sticks and 
  /// the since allocator is also unoptimised so is not likely to have high
  /// performance. It's purpose is to provide space for dynamically written 
  /// thunks, which must have addresses in the range to be 
  /// [imageBase, imageBase + DWORD_MAX] described in the DLL export table.
  /// 
  /// The allocator keeps its data structures external to the allocated 
  /// memory. This allows for locking the page write permissions of the
  /// allocated memory using VirtualProtect which is good security practice
  /// since those regions need to be executable to act as thunks.
  /// </summary>
  class ExternalRegionAllocator
  {
  private:
    struct Block
    {
      unsigned offset;
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
      unsigned _size;
      unsigned _bytesUsed;
      std::set<Block> _blocks;

    public:
      Chunk(unsigned size) : _size(size), _bytesUsed(0) {}

      void* toAddress(void* chunkStart, Block block) const
      {
        return (char*)chunkStart + block.offset;
      }
      Block appendBlock(unsigned blockSize)
      {
        auto[i, success] = _blocks.emplace(Block{ _bytesUsed });
        _bytesUsed += blockSize;
        return *i;
      }
      unsigned getBlockSize(typename decltype(_blocks)::iterator i) const
      {
        assert(i != _blocks.end());
        auto offset = i->offset;
        ++i;
        return i == _blocks.end()
          ? _size - offset
          : i->offset - offset;
      }
      auto findBlock(char* chunkStart, void* memPtr) const
      {
        auto block = Block{ (unsigned)((char*)memPtr - chunkStart) };
        return _blocks.find(block);
      }
      auto splitBlock(char* chunkStart, Block block, unsigned size)
      {
        auto[i, success] = _blocks.emplace(Block{ block.offset + size });
        auto blocksize = getBlockSize(i);
        return std::make_pair(blocksize, FreeBlock{ chunkStart, block.offset + size });
      }
      unsigned free(unsigned size)
      {
        return _bytesUsed -= size;
      }
      unsigned available() const { return _size - _bytesUsed; }
      unsigned size() const { return _size; }
    };

    static constexpr unsigned MIN_BLOCKSIZE = 4;
    static constexpr unsigned GRANULARITY = 16;

  public:
    ExternalRegionAllocator(void* minAddress, void* maxAddress)
      : _minAddress(minAddress)
      , _maxAddress(maxAddress)
    {}
    ~ExternalRegionAllocator()
    {
      for (auto i : _chunks)
        VirtualFree(i.first, 0, MEM_RELEASE);
    }
    auto alloc(unsigned bytesRequested)
    {
      bytesRequested = align<MIN_BLOCKSIZE>(bytesRequested);
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

      auto bytesToAlloc = align<GRANULARITY>(bytesRequested);
      auto* allocated = _allocateChunk(bytesRequested);
      if (!allocated)
        throw _badAllocError;
      _chunks.emplace(std::make_pair((char*)allocated, Chunk(bytesToAlloc)));
      return alloc(bytesRequested);
    }

    auto free(void* memPtr)
    {
      auto iChunk = _chunks.lower_bound((char*)memPtr);
      if (iChunk == _chunks.end() || iChunk->first > memPtr) --iChunk;
      auto [chunkData, chunk] = *iChunk;

      auto iBlock = chunk.findBlock(chunkData, memPtr);
      
      auto blockSize = chunk.getBlockSize(iBlock);

      if (chunk.free(blockSize)== 0)
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
    template<int Tpower>
    inline unsigned align(unsigned val)
    {
      constexpr auto mask = (1 << Tpower) - 1;
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
        address -= 1 << GRANULARITY)
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
    std::bad_alloc _badAllocError;
  };
}