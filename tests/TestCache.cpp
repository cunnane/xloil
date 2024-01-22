#include "CppUnitTest.h"
#include <xloil/ExcelObjCache.h>
#include <chrono>
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::make_unique;
using fmt::format;
using std::vector;
using std::unique_ptr;
using std::string;

namespace Tests
{
  TEST_CLASS(TestCache)
  {
  public:

    TEST_METHOD(LookupCacheTest)
    {
      auto cache = ObjectCache<
        std::unique_ptr<int>,
        CacheUniquifier<std::unique_ptr<int>>>::create();
      const int N = 100;

      ExcelObj caller(L"AnInspector");

      vector<ExcelObj> keys(N);

      for (auto i = 0; i < N; ++i)
        keys[i] = cache->add(make_unique<int>(i), CallerInfo(caller));

      for (auto i = 0; i < N; ++i)
        cache->add(make_unique<int>(i), CallerInfo(caller));

      for (auto i = 0; i < N; ++i)
      {
        auto* val = cache->fetch(keys[i].asStringView());
        Assert::AreEqual<int>(i, **val);
      }
    }

    TEST_METHOD(CallerAddressTypes)
    {
      auto F3 = ExcelObj(msxll::xlref12{ 2, 3, 5, 6 });

      auto sheetName = wstring(L"[Book]Sheet");
      auto caller = CallerInfo(F3, sheetName.c_str());

      Assert::AreEqual(sheetName + L"!R3C6:R4C7", caller.address(AddressStyle::RC | AddressStyle::NOQUOTE));
      Assert::AreEqual(sheetName + L"!F3:G4", caller.address(AddressStyle::A1 | AddressStyle::NOQUOTE));
    }

    TEST_METHOD(CacheV2Test)
    {
      auto cache = ObjectCache<unique_ptr<int>, CacheUniquifierIs<L'X'>>::create();
      const int N = 100;
      vector<ExcelObj> callers;
      vector<ExcelObj> keys(N);

      for (auto i = 0; i < N; ++i)
        callers.emplace_back(ExcelObj(format(L"Key_{0}", i)));

      for (auto i = 0; i < N; ++i)
        keys[i] = cache->add(make_unique<int>(i), CallerInfo(callers[i]));

      for (auto i = 0; i < N; ++i)
      {
        auto* val = cache->fetch(keys[i].asStringView());
        Assert::IsNotNull(val);
        Assert::AreEqual<int>(i, **val);
      }

      vector<ExcelObj> keys2(N);

      cache->onAfterCalculate();

      for (auto i = 0; i < N; ++i)
        keys2[i] = cache->add(make_unique<int>(i), CallerInfo(callers[i]));
  
      auto cacheSize = cache->reap();

      Assert::AreEqual<size_t>(N, cacheSize);

      for (auto i = 0; i < N; ++i)
      {
        auto* val = cache->fetch(keys2[i].asStringView());
        Assert::IsNotNull(val);
        Assert::AreEqual<int>(i, **val);
      }
    }

    TEST_METHOD(CacheSpeedTest1)
    {
      auto& cache = ObjectCacheFactory<std::unique_ptr<int>>::cache();
      const int NumReps = 1;
      const int N = 100;

      vector<ExcelObj> callers;
      vector<ExcelObj> keys(N);
      for (auto i = 0; i < N; ++i)
        callers.emplace_back(ExcelObj(format(L"Key_{0}", i)));
      
      auto t1 = std::chrono::high_resolution_clock::now();

      for (auto i = 0; i < N; ++i)
        keys[i] = cache.add(make_unique<int>(i), CallerInfo(callers[i]));

      for (auto rep = 0; rep < NumReps; ++rep)
        for (auto i = 0; i < N; ++i)
          cache.add(make_unique<int>(i), CallerInfo(callers[i]));

      auto t2 = std::chrono::high_resolution_clock::now();

      for (auto rep = 0; rep < NumReps * 10; ++rep)
        for (auto i = 0; i < N; ++i)
        {
          auto* val = cache.fetch(keys[i].cast<PStringRef>().view());
#ifndef RUN_PERFORMANCE_TEST
          Assert::AreEqual<int>(i, **val);
#endif
        }

#ifndef RUN_PERFORMANCE_TEST
      auto t3 = std::chrono::high_resolution_clock::now();
      auto duration1 = std::chrono::duration_cast<std::chrono::microseconds>(t2 - t1).count();
      auto duration2 = std::chrono::duration_cast<std::chrono::microseconds>(t3 - t2).count();
      Logger::WriteMessage(format("CacheSpeedTest1 - Time 1: {0},   Time 2: {1}", duration1, duration2).c_str());
#endif
    }
  };
}