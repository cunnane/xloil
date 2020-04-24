#include "pch.h"
#include "CppUnitTest.h"
#include <xlOil/ArrayBuilder.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelObj.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;

namespace Tests
{
  TEST_CLASS(ArrayBuilder)
  {
  public:

    TEST_METHOD(ArrayBuild1)
    {
      ExcelArrayBuilder builder(3, 2, 10);
      
      wchar_t* row0[] = { L"Hello", L"World" };
      double row1[] = { 10.0, 7.2 };
      builder.emplace_at(0, 0, row0[0]);
      builder.emplace_at(0, 1, (const wchar_t*)row0[1]);
      builder.emplace_at(1, 0, 10);
      builder.emplace_at(1, 1, 7.2);
      builder.emplace_at(2, 0, CellError::NA);
      builder.setNA(2, 1);

      auto arrayObj = builder.toExcelObj();

      ExcelArray arr(arrayObj);

      Assert::AreEqual<size_t>(arr.nRows(), 2);
      Assert::AreEqual<size_t>(arr.nCols(), 2);
      Assert::AreEqual(arr(0, 0).toString(), wstring(row0[0]));
      Assert::AreEqual(arr(0, 1).toString(), wstring(row0[1]));
      Assert::AreEqual(arr(1, 0).toDouble(), row1[0]);
      Assert::AreEqual(arr(1, 1).toDouble(), row1[1]);
    }
  };
}