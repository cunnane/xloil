#include <xloil/WindowsSlim.h>
#include <stdexcept>
#include <cassert>

namespace xloil
{
  /// <summary>
  /// Allocator which returns memory regions in a specified address range.
  /// This is done by repeated calls to VirtualAlloc and the since allocator 
  /// is also unoptimised so is not likely to be performant. It's purpose is
  /// to provide space for dynamically written thunks, which must have 
  /// addresses in the range to be [imageBase, imageBase + DWORD_MAX]
  /// described in the DLL export table.
  /// 
  /// The allocator is inspired by the simple example here:
  /// https://devblogs.microsoft.com/oldnewthing/20050519-00/?p=35603
  /// but has grown unnecessarily complex.
  /// </summary>
  template <int TMemFlags = 0, int TPageAccess = PAGE_READWRITE>
  class RegionAllocator
  {
  private:

    struct ChunkHeader
    {
      ChunkHeader* prev;
      unsigned  size;
      unsigned  bytesUsed;
    };
    struct BlockHeader
    {
      unsigned size;
    };
    struct FreeListItem : public BlockHeader
    {
      FreeListItem* next;
    };

    static constexpr unsigned MIN_BLOCKSIZE = 4;
    static constexpr unsigned GRANULARITY = 16;

  public:
    RegionAllocator(void* minAddress, void* maxAddress)
      : _nextByte(nullptr), _endOfChunk(nullptr), _currentChunk(nullptr)
      , _freeList(nullptr)
      , _minAddress(minAddress)
      , _maxAddress(maxAddress)
    {
      static_assert(1 << MIN_BLOCKSIZE == sizeof(FreeListItem));
    }

    ~RegionAllocator()
    {
      auto* phdr = _currentChunk;
      while (phdr) 
      {
        ChunkHeader hdr = *phdr;
        VirtualFree(hdr.prev, 0, MEM_RELEASE);
        phdr = hdr.prev;
      }
    }

    auto alloc(unsigned bytesRequested)
    {
      bytesRequested = align<MIN_BLOCKSIZE>(bytesRequested);
      auto bytesRequired = (unsigned)sizeof(BlockHeader) + bytesRequested;

      if (_nextByte + bytesRequired <= _endOfChunk)
      {
        auto region = (BYTE*)new (_nextByte) BlockHeader{ bytesRequested } + sizeof(BlockHeader);
        _nextByte = _nextByte + bytesRequired;
        _currentChunk->bytesUsed += bytesRequired;
        return region;
      }

      // Search free list
      auto thisFree = _freeList;
      for (auto prevFree = thisFree; thisFree != nullptr; prevFree = thisFree, thisFree = thisFree->next)
      {
        if (bytesRequested <= thisFree->size)
        {
          removeFromFreeList(thisFree, prevFree);
          
          if (thisFree->size - bytesRequested > 1 << MIN_BLOCKSIZE)
          {
            // Split the block
            thisFree->size -= bytesRequested;
            auto newFreeItem = reinterpret_cast<FreeListItem*>((BYTE*)thisFree + bytesRequired);
            newFreeItem->next = nullptr;
            newFreeItem->size = thisFree->size - bytesRequired;
            thisFree->size = bytesRequested;
            appendToFreeList(newFreeItem);
          }
          
          auto* chunkHeader = getChunkHeader(thisFree);
          chunkHeader->bytesUsed += bytesRequired;
          return (BYTE*)thisFree + sizeof(BlockHeader);
        }
      }

      auto bytesToAlloc = align<GRANULARITY>(bytesRequired + (unsigned)sizeof(ChunkHeader));
      auto allocated = (BYTE*)_allocateChunk(bytesToAlloc);
      if (!allocated)
        throw _badAlloc;

      _endOfChunk = allocated + bytesToAlloc;
      _currentChunk = new (allocated) ChunkHeader{ _currentChunk, bytesToAlloc, 0 };
      _nextByte = allocated + sizeof(ChunkHeader);
      return this->alloc(bytesRequested);
    }

    auto free(void* p)
    {
      auto* blockHeader = reinterpret_cast<BlockHeader*>(p) - 1;
      auto* chunkHeader = getChunkHeader(p);

      // Decrement use count in current chunk
      chunkHeader->bytesUsed -= blockHeader->size + sizeof(BlockHeader);

      // Never free current chunk
      if (chunkHeader->bytesUsed == 0 && chunkHeader != _currentChunk)
      {
        // Need to remove all freelist entries in this chunk
        auto* nextFree = _freeList;
        void* chunkEnd =  chunkHeader + chunkHeader->size;
        for (auto prevFree = nextFree; nextFree != nullptr; prevFree = nextFree, nextFree = nextFree->next)
        {
          if ((void*)chunkHeader <= nextFree && nextFree < chunkEnd)
            removeFromFreeList(nextFree, prevFree);
        }
        auto* next = nextChunk(chunkHeader);
        next->prev = chunkHeader->prev;
        VirtualFree(chunkHeader, 0, MEM_RELEASE);
      }
      else
      {
        auto* freeItem = reinterpret_cast<FreeListItem*>(blockHeader);
        freeItem->next = nullptr;
        appendToFreeList(freeItem);
      }
    }

    void* currentPage() const
    {
      return pageStart(_nextByte);
    }

  private:
    void* pageStart(void* memPtr) const
    {
      return (void*)((size_t)memPtr & ~0xFFFF);
    }

    ChunkHeader* getChunkHeader(void* memPtr) const
    {
      auto page = pageStart(memPtr);
      auto p = _currentChunk;
      while (p > page) p = p->prev;
      return p;
    }

    FreeListItem* freeListTail() const
    {
      auto p = _freeList;
      if (p)
        for (; p->next != nullptr; p = p->next);
      return p;
    }

    void appendToFreeList(FreeListItem* item)
    {
      auto tail = freeListTail();
      if (tail)
        tail->next = item;
      else
        _freeList = item;
    }

    void removeFromFreeList(FreeListItem* item, FreeListItem* prev)
    {
      if (_freeList == item)
        _freeList = item->next;
      else
        prev->next = item->next;
    }
    ChunkHeader* nextChunk(ChunkHeader* chunk) const
    {
      // Assume this eventually terminates
      auto p = _currentChunk;
      for (; p->prev != chunk; p = p->prev);
      return p;
    }

    inline void* _allocateChunk(unsigned numBytes)
    {
      void* p = nullptr;
      for (auto address = (BYTE*)(_currentChunk ? _currentChunk : _maxAddress); 
          address > _minAddress && p == nullptr; 
          address -= 1 << GRANULARITY)
        p = VirtualAlloc(address, numBytes, MEM_COMMIT | MEM_RESERVE, TPageAccess);
      return p;
    }

    template<int Tpower>
    inline size_t align(size_t val)
    {
      constexpr auto mask = (1 << Tpower) - 1;
      return val + mask & ~(mask);
    }
    template<int Tpower>
    inline unsigned align(unsigned val)
    {
      constexpr auto mask = (1 << Tpower) - 1;
      return val + mask & ~(mask);
    }
  private:
    BYTE* _endOfChunk;
    BYTE* _nextByte;
    ChunkHeader* _currentChunk;
    std::bad_alloc _badAlloc;
    FreeListItem* _freeList;
    void* _maxAddress;
    void* _minAddress;
  };
}