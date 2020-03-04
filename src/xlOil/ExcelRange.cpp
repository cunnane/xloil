#include "ExcelRange.h"

namespace xloil
{
  XLOIL_EXPORT ExcelRange::ExcelRange(const ExcelObj& from)
  {
    switch (from.type())
    {
    case ExcelType::SRef:
    {
      callExcelRaw(msxll::xlSheetId, this);
      auto r = from.val.sref.ref;
      create(val.mref.idSheet, r.rwFirst, r.colFirst, r.rwLast, r.colLast);
      break;
    }
    case ExcelType::Ref:
      val.mref.idSheet = from.val.mref.idSheet;
      val.mref.lpmref = from.val.mref.lpmref;
      xltype = from.xltype;
      if (val.mref.lpmref->count != 1)
        XLO_THROW("Only contiguous refs supported");
      break;
    default:
      XLO_THROW("ExcelRange: expecting reference type");
    }
  }

  XLOIL_EXPORT ExcelRange::ExcelRange(const wchar_t* address)
    : ExcelRange(callExcel(msxll::xlfIndirect, address))
  {
  }

  XLOIL_EXPORT ExcelRange::ExcelRange(
    msxll::IDSHEET sheetId, int fromRow, int fromCol, int toRow, int toCol)
  {
    create(sheetId, fromRow, fromCol, toRow, toCol);
  }

  XLOIL_EXPORT std::wstring ExcelRange::address(bool local) const
  {
    auto& ref = this->val.mref.lpmref->reftbl[0];
    if (local)
      return captureWinApiString([ref](auto buf, auto sz)
    {
      return xlrefToStringA1(ref, buf, sz);
    },
        CELL_ADDRESS_A1_MAX);
    else
      return captureWinApiString([this, ref](auto buf, auto sz)
    {
      return xlrefSheetAddressA1(this->sheetId(), ref, buf, sz, true);
    });
  }

  void ExcelRange::create(
    msxll::IDSHEET sheetId, int fromRow, int fromCol, int toRow, int toCol)
  {
    xltype = msxll::xltypeRef | msxll::xlbitDLLFree;
    val.mref.idSheet = sheetId;
    val.mref.lpmref = new msxll::XLMREF12[1];
    val.mref.lpmref->count = 1;
    ref().rwFirst = fromRow;
    ref().colFirst = fromCol;
    ref().rwLast = toRow;
    ref().colLast = toCol;
    if (fromRow > toRow)
      XLO_THROW("Empty range: fromRow={0}, toRow={1}, ", fromRow, toRow);
    if (fromCol > toCol)
      XLO_THROW("Empty range: fromCol={0}, toCol={1}", fromCol, toCol);
  }
}