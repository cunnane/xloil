#include "CppUnitTest.h"
#include <xloil-Dynamic/RegionAllocator.h>
#include <vector>
#include <list>
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::string;
using std::vector;
using std::list;

namespace Tests
{
  TEST_CLASS(TestSimpleAllocator)
  {
  public:
    TEST_METHOD(AllocatorTest1)
    {
      //SYSTEM_INFO si;
      //GetSystemInfo(&si);
      //assert(si.dwAllocationGranularity == m_dwGranularity);
      //_maxAddress = si.lpMaximumApplicationAddress;

      SYSTEM_INFO si;
      GetSystemInfo(&si);
        
      auto allocator = RegionAllocator(si.lpMinimumApplicationAddress, si.lpMaximumApplicationAddress);

      vector<void*> ptrs;
      for (auto sz = 16; sz < 1024; sz += 32)
      {
        ptrs.push_back(allocator.alloc(sz));
      }
      for (auto p : ptrs)
        allocator.free(p);
    }

    TEST_METHOD(AllocatorTest2)
    {
      SYSTEM_INFO si;
      GetSystemInfo(&si);

      auto allocator = RegionAllocator(si.lpMinimumApplicationAddress, si.lpMaximumApplicationAddress);

      const char* sample = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor";
      auto sampleLen = strlen(sample);
      vector<char*> ptrs;
      for (auto i = 1; i < sampleLen; ++i)
      {
        auto str = (char*)allocator.alloc(i + 1);
        strncpy_s(str, i + 1, sample, i);
        ptrs.push_back(str);
      }
      for (auto i = 0; i < ptrs.size(); i += 4)
      {
        allocator.free(ptrs[i]);
        ptrs[i] = nullptr;
      }
      for (auto i = 1; i < sampleLen; ++i)
      {
        auto str = (char*)allocator.alloc(i + 1);
        strncpy_s(str, i + 1, sample, i);
        ptrs.push_back(str);
      }
      for (auto str : ptrs)
      {
        if (str)
          Assert::AreEqual(0, strncmp(str, sample, strlen(str)));
      }
    }
  };
}