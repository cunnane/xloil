#include "COMRange.h"
#include <xlOil/ExcelObj.h>
#include "Connect.h"
#include "ExcelTypeLib.h"
#include "ComVariant.h"
#include <xlOil/ExcelCall.h>


namespace xloil
{
  namespace COM
  {
    ExcelRange rangeFromAddress(const wchar_t* address)
    {
      try
      {
        auto& app = excelApp();
        auto range = app.GetRange(_variant_t(address));
        return rangeFromComRange(range);
      }
      catch (_com_error& error)
      {
        XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
      }
    }

    void rangeSetValue(ExcelRange& range, const ExcelObj& value)
    {
      try
      {
        auto& app = excelApp();
        const auto address = range.address();
        const auto cRange = app.GetRange(_variant_t(address.c_str()));
        cRange->Value2 = excelObjToVariant(value);
      }
      catch (_com_error& error)
      {
        XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
      }
    }

    ExcelRange rangeFromComRange(Excel::Range* range)
    {
      try
      {
        const auto nCols = range->Columns->Count;
        const auto nRows = range->Rows->Count;

        // Excel uses 1-based indexing for these, so we adjust
        const auto fromRow = range->Row - 1;
        const auto fromCol = range->Column - 1;

        // Convert to an XLL SheetId
        auto wb = (Excel::_WorkbookPtr)range->Worksheet->Parent;
        const auto sheetId =
          callExcel(msxll::xlSheetId, fmt::format(L"[{0}]{1}",
            wb->Name, range->Worksheet->Name));

        return ExcelRange(sheetId.val.mref.idSheet,
          fromRow, fromCol, fromRow + nRows - 1, fromCol + nCols - 1);
      }
      catch (_com_error& error)
      {
        XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
      }
    }
  }
}