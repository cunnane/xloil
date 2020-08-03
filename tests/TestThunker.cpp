#include "CppUnitTest.h"
#include <xloilHelpers/Thunker.h>
#include <xloil/Register.h>
#include <xloil/ExcelObj.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;

namespace Tests
{
  ExcelObj* callback(void* data, const ExcelObj** p) noexcept
  {
    auto n = *(int*)data;
    return (ExcelObj*)p[n];
  }
  void asyncCallback(void* data, const ExcelObj* r, const ExcelObj** p) noexcept
  {
    auto n = *(int*)data;
    *(ExcelObj*)r = *p[n];
  }
  TEST_CLASS(ThunkerTests)
  {
  public:

    TEST_METHOD(TestHandRoll)
    {
      // The context selects the argument to return in the callbacks
      int context = 0;
      const auto pContext = (void*)&context;

      constexpr auto bufSize = 256u;
      char buffer1[bufSize]; // Used for asmjit compiled thunk
      char buffer2[bufSize]; // Used for hand rolled thunk
      size_t codeSize;
      
      buildThunk    (callback, pContext, 2, buffer1, bufSize, codeSize);
      buildThunkLite(callback, pContext, 2, buffer2, bufSize, codeSize);

      ExcelObj arg1(7);
      ExcelObj arg2(3);

      {
        context = 0;

        typedef ExcelObj* (*TwoArgs)(ExcelObj*, ExcelObj*);
        auto result1 = ((TwoArgs)(void*)buffer1)(&arg1, &arg2);
        auto result2 = ((TwoArgs)(void*)buffer2)(&arg1, &arg2);

        Assert::IsTrue(*result1 == arg1);
        Assert::IsTrue(*result2 == arg1);
      }

      buildThunk    (callback, pContext, 7, buffer1, bufSize, codeSize);
      buildThunkLite(callback, pContext, 7, buffer2, bufSize, codeSize);

      {
        context = 5;

        typedef ExcelObj* (*SevenArgs)(ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*);
        auto result1 = ((SevenArgs)(void*)buffer1)(&arg1, &arg1, &arg1, &arg1, &arg1, &arg2, &arg1);
        auto result2 = ((SevenArgs)(void*)buffer2)(&arg1, &arg1, &arg1, &arg1, &arg1, &arg2, &arg1);

        Assert::IsTrue(*result1 == arg2);
        Assert::IsTrue(*result2 == arg2);
      }

      buildThunk    (asyncCallback, pContext, 3, buffer1, bufSize, codeSize);
      buildThunkLite(asyncCallback, pContext, 3, buffer2, bufSize, codeSize);

      {
        context = 1;
        ExcelObj asyncReturn;

        typedef void (*ThreeArgsAsync)(ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*);

        ((ThreeArgsAsync)(void*)buffer1)(&asyncReturn, &arg2, &arg1, &arg2);
        Assert::IsTrue(asyncReturn == arg1);

        ((ThreeArgsAsync)(void*)buffer2)(&asyncReturn, &arg2, &arg1, &arg2);
        Assert::IsTrue(asyncReturn == arg1);
      }
    }
  };
}