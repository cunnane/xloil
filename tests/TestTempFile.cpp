#include "CppUnitTest.h"
#include <xloilHelpers/Environment.h>
#include <xloil/WindowsSlim.h>
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::unique_ptr;

namespace Tests
{
  TEST_CLASS(TempFileTests)
  {
  public:
    TEST_METHOD(TestTempFile1)
    {
      // TODO: TestTempFile1 is a little too rudimentary!
      auto [handle, name] = Helpers::makeTempFile();
      CloseHandle(handle);
    }
  };
}