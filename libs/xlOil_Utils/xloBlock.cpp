#include "ExcelObj.h"
#include "ArrayBuilder.h"
#include "ExcelArray.h"
#include <xloil/StaticRegister.h>
#include <boost/preprocessor/repeat_from_to.hpp>

using std::vector;
using std::tuple;

/*
Layout should be
"1, 2, 3 ; 4, 5 ; 6"
"1, 2, 3 ; 4, , 5"
TODO: support transpose with a t
*/
namespace xloil
{
#define XLOBLOCK_NARGS 20
#define DECLARE_ARG(z,N,d) , const ExcelObj& arg##N
#define COPY_ARG(z,N,d) args[N - 1] = &arg##N;

  XLO_FUNC xloBlock(const ExcelObj& layout
    BOOST_PP_REPEAT_FROM_TO(1, XLOBLOCK_NARGS, DECLARE_ARG, data)
  )
  {
    try
    {
      constexpr size_t nArgs = XLOBLOCK_NARGS - 1;
      const ExcelObj* args[nArgs];
      BOOST_PP_REPEAT_FROM_TO(1, XLOBLOCK_NARGS, COPY_ARG, data)

      constexpr int NEW_ROW = -1;
      constexpr int PADDING = 0;

      const auto parse = layout.toString();

      if (parse.empty())
        XLO_THROW("No layout");

      vector<tuple<int, size_t, size_t>> spec;

      auto* p = parse.data();

      size_t totalRows = 0, totalCols = 0;
      size_t colsForCurrent = 0, rowsForCurrent = 0;
      size_t totalStrLength = 0;
      while (true)
      {
        wchar_t* nextP;
        auto argNum = wcstol(p, &nextP, 10);
        if (argNum > nArgs)
          XLO_THROW("Arg {i} out of range", argNum);
        p = nextP;

        if (argNum == 0)
          spec.emplace_back(PADDING, 0, 0);
        else
        {
          size_t nRows = 1, nCols = 1;
          auto* obj = args[argNum - 1];
          switch (obj->type())
          {
          case ExcelType::Multi:
          {
            obj->trimmedArraySize(nRows, nCols);
            auto arrData = obj->asArray();
            auto endArray = arrData + nRows * nCols;
            for (; arrData != endArray; ++arrData)
              totalStrLength += arrData->stringLength();
          }
          case ExcelType::Str:
            totalStrLength += obj->stringLength();
            break;
          default:
            break;
          }
          spec.emplace_back(argNum, nRows, nCols);
          rowsForCurrent = std::max(rowsForCurrent, nRows);
          colsForCurrent += nCols;
        }
        auto separator = wcspbrk(p, L",;");

        if (!separator || *separator == L';')
        {
          spec.emplace_back(NEW_ROW, rowsForCurrent, 0);
          totalCols = std::max(totalCols, colsForCurrent);
          totalRows += rowsForCurrent;
          colsForCurrent = 0;
          rowsForCurrent = 0;
        }

        if (!separator)
          break;

        p = separator + 1;
      }

      if (totalCols == 0 || totalRows == 0)
        return ExcelObj::returnValue(CellError::NA);

      // TODO: avoid looping the spec twice!
      size_t row = 0, col = 0;
      for (auto i = 0; i < spec.size(); ++i)
      {
        auto[iArg, nRows, nCols] = spec[i];
        switch (iArg)
        {
        case NEW_ROW:
          col = 0;
          break;
        case PADDING:
        {
          auto j = i;
          while (std::get<0>(spec[++j]) != NEW_ROW)
            col += std::get<2>(spec[j]);
          spec[i] = std::make_tuple(PADDING, 0, totalCols - col);
          i = j;
          col = 0;
          break;
        }
        default:
          col += nCols;
        }
      }

      ExcelArrayBuilder builder(totalRows, totalCols, totalStrLength);

      // TODO: optional fill specifier?
      for (auto i = 0; i < totalRows; ++i)
        for (auto j = 0; j < totalCols; ++j)
          builder.setNA(i, j);

      row = 0; col = 0;
      for (auto[iArg, nRows, nCols] : spec)
      {
        switch (iArg)
        {
        case NEW_ROW:
          col = 0;
          row += nRows;
          break;
        case PADDING:
          col += nCols;
          break;
        default:
        {
          auto* obj = args[iArg - 1];
          switch (obj->type())
          {
          case ExcelType::Multi:
          {
            ExcelArray arr(*obj, nRows, nCols);

            for (auto i = 0; i < nRows; ++i)
              for (auto j = 0; j < nCols; ++j)
                builder.emplace_at(row + i, col + j, arr(i, j));

            col += nCols;
            break;
          }
          default:
            builder.emplace_at(row, col, *obj);
            ++col;
          }
        }
        }
      }

      return ExcelObj::returnValue(builder.toExcelObj());
    }
    catch (const std::exception& e)
    {
      XLO_RETURN_ERROR(e);
    }
  }
  XLO_REGISTER(xloBlock).threadsafe()
    .help(L"Creates a block matrix given a layout specification."
      "Layout has the form '1, 2, 3; 4, ,5' where numbers "
      "refer to the n-th argument (1-based). Semi-colon indicates"
      "a new row. Two consecutive commas mean an automatically "
      "resized padded space. Whitespace is ignored.")
    .arg(L"Layout", L"String specifying how to layout the blocks");
}