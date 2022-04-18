#include <xloil/ExcelRange.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/ExcelRef.h>
#include <xlOil/AppObjects.h>
#include <xlOil-COM/XllContextInvoke.h>
#include <xlOil-COM/ComVariant.h>

namespace xloil
{
  namespace
  {
    _variant_t stringToVariant(const std::wstring_view& str)
    {
      auto variant = COM::stringToVariant(str);
      return _variant_t(variant, false);
    }
  }
  Range* newRange(const wchar_t* address)
  {
    if (InXllContext::check())
      return new XllRange(ExcelRef(address));
    else
      return new ExcelRange(address);
  }
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
        callExcel(msxll::xlSheetId, XLO_FMT(L"[{0}]{1}",
          (const char*)wb->Name, (const char*)range->Worksheet->Name));

      return ExcelRef(sheetId.val.mref.idSheet,
        fromRow, fromCol, fromRow + nRows - 1, fromCol + nCols - 1);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange::ExcelRange(const std::wstring_view& address)
  {
    try
    {
      auto addressStr = COM::stringToVariant(address);
      auto rangePtr = excelApp().GetRange(_variant_t(addressStr, false));
      init(rangePtr);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange::ExcelRange(const Range& range)
  {
    auto excelRange = dynamic_cast<const ExcelRange*>(&range);
    if (excelRange)
      init(excelRange->ptr());
    else
    {
      init(nullptr);
      *this = ExcelRange(range.address());
    }
  }

  Range* ExcelRange::range(
    int fromRow, int fromCol,
    int toRow, int toCol) const
  {
    try
    {
      if (toRow == Range::TO_END)
        toRow = ptr()->Row + ptr()->Rows->GetCount();
      if (toCol == Range::TO_END)
        toCol = ptr()->Column + ptr()->Columns->GetCount();

      // Caling range->GetRange(cell1, cell2) does a very weird thing
      // which I can't make sense of. Better to call ws.range(...)
      auto ws = (Excel::_WorksheetPtr)ptr()->Parent;
      auto r = ws->GetRange(
        ptr()->Cells->Item[fromRow + 1][fromCol + 1],
        ptr()->Cells->Item[toRow + 1][toCol + 1]);
      return new ExcelRange(r);
    }
    XLO_RETHROW_COM_ERROR;
  }

  std::tuple<Range::row_t, Range::col_t> ExcelRange::shape() const
  {
    return { ptr()->Rows->GetCount(), ptr()->Columns->GetCount() };
  }

  std::tuple<Range::row_t, Range::col_t, Range::row_t, Range::col_t> ExcelRange::bounds() const
  {
    return {
      ptr()->Row,
      ptr()->Column,
      ptr()->Row + ptr()->Rows->GetCount() - 1,
      ptr()->Column + ptr()->Columns->GetCount() - 1
    };
  }

  std::wstring ExcelRange::address(bool local) const
  {
    try
    {
      auto result = local
        ? ptr()->GetAddress(true, true, Excel::xlA1)
        : ptr()->GetAddressLocal(true, true, Excel::xlA1);
      return std::wstring(result);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelObj ExcelRange::value() const
  {
    return refFromComRange(ptr()).value();
  }

  ExcelObj ExcelRange::value(row_t i, col_t j) const
  {
    return COM::variantToExcelObj(ptr()->Cells->Item[i + 1][j + 1]);
  }

  void ExcelRange::set(const ExcelObj& value)
  {
    try
    {
      VARIANT v;
      COM::excelObjToVariant(&v, value);
      ptr()->PutValue2(_variant_t(v, false)); // Move variant
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelRange::setFormula(const std::wstring_view& formula)
  {
    try
    {
      if (size() > 1)
        ptr()->FormulaArray = stringToVariant(formula);
      else
        ptr()->Formula = stringToVariant(formula);
    }
    XLO_RETHROW_COM_ERROR;
  }

  std::wstring ExcelRange::formula()
  {
    try
    {
      return ((_bstr_t)ptr()->Formula).GetBSTR();
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelRange::clear()
  {
    ptr()->Clear();
  }

  std::wstring ExcelRange::name() const
  {
    return address(false);
  }

  ExcelWorksheet ExcelRange::parent() const
  {
    try
    {
      return ExcelWorksheet(ptr()->Worksheet);
    }
    XLO_RETHROW_COM_ERROR;
  }
}