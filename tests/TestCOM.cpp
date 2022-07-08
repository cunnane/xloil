#include "CppUnitTest.h"
#include <xlOil/AppObjects.h>
#include <xloil/ExcelTypeLib.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::unique_ptr;

namespace Tests
{
  TEST_CLASS(TestCOM)
  {
  public:
    TEST_METHOD(TestCOM1)
    {
      // Just a smoke test at the moment, but hey it's better than nowt
      auto app = Application();
      app.workbooks();
      app.quit();
    }
  };
}