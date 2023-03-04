#include "CppUnitTest.h"
#include <xloil-Dynamic/Thunker.h>
#include <xloil/Register.h>
#include <xloil/ExcelObj.h>
#include <xloil/WindowsSlim.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::unique_ptr;

namespace Tests
{
  ExcelObj* callback(void* data, const ExcelObj** p) noexcept
  {
    auto n = *(int*)data;
    return (ExcelObj*)p[n];
  }
  void asyncCallback(void* data, const ExcelObj** p) noexcept
  {
    auto n = *(int*)data;
    const ExcelObj* r = p[0];
    *(ExcelObj*)r = *p[n + 1];
  }
  TEST_CLASS(ThunkerTests)
  {
  public:

    TEST_METHOD(PatchThunk)
    {
      auto callback = (void*)0xABBAABBA;
      auto context = (void*)0xABFAB;
      char code[1024];

      ThunkWriter writer(callback, context, 3, true);
      
      auto codeSize = writer.writeCode(code, _countof(code));

      auto newContext = (void*)0xBABABA;
      auto success = patchThunkData(code, codeSize, context, newContext);
      Assert::IsTrue(success);
    }

#ifdef _WIN64
    TEST_METHOD(TestHandRoll)
    {
      // The context selects the argument to return in the callbacks
      int context = 0;
      const auto pContext = (void*)&context;

      constexpr auto bufSize = 256u;
      char buffer1[bufSize]; // Used for asmjit compiled thunk
      char buffer2[bufSize]; // Used for hand rolled thunk
      
      DWORD dummy;
      Assert::IsTrue(VirtualProtect(buffer1, bufSize, PAGE_EXECUTE_READWRITE, &dummy));
      Assert::IsTrue(VirtualProtect(buffer2, bufSize, PAGE_EXECUTE_READWRITE, &dummy));

      ThunkWriter(callback, pContext, 2, false).writeCode(buffer1, bufSize);
      ThunkWriter(callback, pContext, 2, false, ThunkWriter::SLOW_BUILD).writeCode(buffer2, bufSize);

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

      ThunkWriter(callback, pContext, 7, true).writeCode(buffer1, bufSize);
      ThunkWriter(callback, pContext, 7, true, ThunkWriter::SLOW_BUILD).writeCode(buffer2, bufSize);

      {
        context = 5;

        typedef ExcelObj* (*SevenArgs)(ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*, ExcelObj*);
        auto result1 = ((SevenArgs)(void*)buffer1)(&arg1, &arg1, &arg1, &arg1, &arg1, &arg2, &arg1);
        auto result2 = ((SevenArgs)(void*)buffer2)(&arg1, &arg1, &arg1, &arg1, &arg1, &arg2, &arg1);

        Assert::IsTrue(*result1 == arg2);
        Assert::IsTrue(*result2 == arg2);
      }

      ThunkWriter(asyncCallback, pContext, 3, false).writeCode(buffer1, bufSize);
      ThunkWriter(asyncCallback, pContext, 3, false, ThunkWriter::SLOW_BUILD).writeCode(buffer2, bufSize);

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
#endif
  };
}