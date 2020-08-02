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
      builder(0, 0) = row0[0];
      builder(0, 1) = (const wchar_t*)row0[1];
      builder(1, 0) = 10;
      builder(1, 1) = 7.2;
      builder(2, 0) = CellError::NA;
      builder(2, 1) = CellError::NA;

      auto arrayObj = builder.toExcelObj();

      ExcelArray arr(arrayObj);

      Assert::AreEqual<size_t>(arr.nRows(), 2);
      Assert::AreEqual<size_t>(arr.nCols(), 2);
      Assert::AreEqual(arr(0, 0).toString(), wstring(row0[0]));
      Assert::AreEqual(arr(0, 1).toString(), wstring(row0[1]));
      Assert::AreEqual(arr(1, 0).toDouble(), row1[0]);
      Assert::AreEqual(arr(1, 1).toDouble(), row1[1]);
    }

    TEST_METHOD(ArrayAccess)
    {
      ExcelArrayBuilder builder(6, 4);
      builder.fillNA();

      for (auto i = 0u; i < builder.nRows(); ++i)
        for (auto j = 0u; j < builder.nCols(); ++j)
          builder(i, j) = i * j;

      ExcelArray array(builder.toExcelObj());
      Assert::AreEqual(array.nRows(), builder.nRows());
      Assert::AreEqual(array.nCols(), builder.nCols());

      for (auto n = 1u; n < array.nCols(); ++n)
      {
        auto sub = array.subArray(0, 1, -1, n);

        Assert::AreEqual(array.nRows(), sub.nRows());

        // Check sub-array equals array using a whole-array index
        auto k = 0u;
        for (auto i = 0u; i < sub.nRows(); ++i)
          for (auto j = 1u; j < n; ++j, ++k)
            Assert::IsTrue(array(i, j) == sub(k));

        // Check sub-array whole-index matches iterator
        k = 0;
        for (auto& val : sub)
          Assert::IsTrue(val == sub(k++));
      }
    }
  };
}