#include <xloil/ExcelRange.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/ExcelRef.h>
#include <xlOil/AppObjects.h>
#include <xlOil-COM/XllContextInvoke.h>
#include <xlOil-COM/ComVariant.h>

namespace xloil
{
  Range* newRange(const wchar_t* address)
  {
    if (InXllContext::check())
      return new XllRange(ExcelRef(address));
    else
      return new ExcelRange(address);
  }

  namespace
  {
    ExcelRef refFromComRange(Excel::Range* range)
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

        return ExcelRef(sheetId.val.mref.idSheet,
          fromRow, fromCol, fromRow + nRows - 1, fromCol + nCols - 1);
      }
      catch (_com_error& error)
      {
        XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
      }
    }
  }

  ExcelRange::ExcelRange(const wchar_t* address)
  {
    try
    {
      auto& app = excelApp();
      _range = app.GetRange(_variant_t(address));
    }
    catch (_com_error& error)
    {
      XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
    }
  }
  ExcelRange::ExcelRange(Excel::Range* range)
    : _range(range)
  {
    range->AddRef();
  }

  Range* ExcelRange::range(
    int fromRow, int fromCol,
    int toRow, int toCol) const
  {
    if (toRow == Range::TO_END)
      toRow = _range->Row + _range->Rows->GetCount();
    if (toCol == Range::TO_END)
      toCol = _range->Column + _range->Columns->GetCount();

    auto r = _range->GetRange(
      _range->Cells->Item[fromRow - 1][fromCol - 1],
      _range->Cells->Item[toRow - 1][toCol - 1]);
    return new ExcelRange(r);
  }


  std::tuple<Range::row_t, Range::col_t> ExcelRange::shape() const
  {
    return { _range->Rows->GetCount(), _range->Columns->GetCount() };
  }

  std::tuple<Range::row_t, Range::col_t, Range::row_t, Range::col_t> ExcelRange::bounds() const
  {
    return {
      _range->Row,
      _range->Column,
      _range->Row + _range->Rows->GetCount() - 1,
      _range->Column + _range->Columns->GetCount() - 1
    };
  }

  std::wstring ExcelRange::address(bool local) const
  {
    try
    {
      auto result = local
        ? _range->GetAddress(true, true, Excel::xlA1)
        : _range->GetAddressLocal(true, true, Excel::xlA1);
      return std::wstring(result);
    }
    catch (_com_error& error)
    {
      XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
    }
  }

  ExcelObj ExcelRange::value() const
  {
    // TODO: not the best implementation. May even crash Excel.
    // On the other hand, it's not ideal to marshal COM -> ExcelObj -> Other language
    return refFromComRange(_range).value();
  }

  ExcelObj ExcelRange::value(row_t i, col_t j) const
  {
    // TODO: not the best implementation. May even crash Excel.
    return refFromComRange(Excel::RangePtr(_range->Cells->Item[i][j])).value();
  }

  void ExcelRange::set(const ExcelObj& value)
  {
    try
    {
      VARIANT v;
      COM::excelObjToVariant(&v, value);
      _range->PutValue2(_variant_t(v, false)); // Move variant
    }
    catch (_com_error& error)
    {
      XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
    }
  }
  void ExcelRange::clear()
  {
    _range->Clear();
  }

  std::wstring ExcelRange::name() const
  {
    return address(false);
  }
}