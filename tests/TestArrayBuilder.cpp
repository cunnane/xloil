#include "CppUnitTest.h"
#include <xlOil/ArrayBuilder.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelObj.h>
#include <vector>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::vector;

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
      Assert::AreEqual(arr(1, 0).get<double>(), row1[0]);
      Assert::AreEqual(arr(1, 1).get<double>(), row1[1]);
    }

    TEST_METHOD(ArrayIteratorSyntax)
    {
      vector<double> data = { 1, 2, 3, 4, 5, 6 };
      ExcelArrayBuilder builder(3, 2);
      std::copy(data.begin(), data.end(), builder.begin());

      auto arrayData = builder.toExcelObj();
      ExcelArray array(arrayData);

      for (auto i = 0u; i < data.size(); ++i)
        Assert::AreEqual(array[i].get<double>(), data[i]);
    }

    TEST_METHOD(ArrayAccess)
    {
      ExcelArrayBuilder builder(6, 4);
      builder.fillNA();

      for (auto i = 0u; i < builder.nRows(); ++i)
        for (auto j = 0u; j < builder.nCols(); ++j)
          builder(i, j) = i * j;

      auto arrayData = builder.toExcelObj();
      ExcelArray array(arrayData);

      Assert::AreEqual(array.nRows(), builder.nRows());
      Assert::AreEqual(array.nCols(), builder.nCols());

      for (auto n = 1u; n < array.nCols(); ++n)
      {
        auto sub = array.slice(0, 1, array.nRows(), n);

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

    TEST_METHOD(ArrayIterators)
    {
      ExcelArrayBuilder builder(6, 4);
      builder.fillNA();

      for (auto i = 0u; i < builder.nRows(); ++i)
        for (auto j = 0u; j < builder.nCols(); ++j)
          builder(i, j) = i * j;

      auto arrayData = builder.toExcelObj();
      ExcelArray array(arrayData);

      for (auto iCol = 0; iCol < array.nCols(); ++iCol)
      {
        auto iRow = 0;
        for (auto p = array.col_begin(iCol); p != array.col_end(iCol); ++p, ++iRow)
        {
          Assert::AreEqual(iCol * iRow, p->get<int>());
        }
      }
    }

    TEST_METHOD(SubArrayAccess)
    {
      constexpr auto R = 6, C = 4;
      ExcelArrayBuilder builder(R, C);
      builder.fillNA();

      for (auto i = 0u; i < builder.nRows(); ++i)
        for (auto j = 0u; j < builder.nCols(); ++j)
          builder(i, j) = i * j;

      auto arrayData = builder.toExcelObj();
      ExcelArray array(arrayData);

      for (auto n = -R + 1; n < R - 1; ++n)
      {
        auto sub = array.slice(n, 1, R, 2);
        Assert::IsTrue(sub(0) == array(n < 0 ? R + n : n, 1));
      }
    }
  };
}