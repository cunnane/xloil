#include "ExcelRange.h"
#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelState.h>
#include <xloilHelpers/StringUtils.h>
#include <xloilHelpers/Environment.h>
#include <Cominterface/XllContextInvoke.h>
#include <Cominterface/COMRange.h>

namespace xloil
{
  XLOIL_EXPORT ExcelRange::ExcelRange(const ExcelObj& from)
  {
    switch (from.type())
    {
    case ExcelType::SRef:
    {
      callExcelRaw(msxll::xlSheetId, this); // TODO: may not work as expected in macro funcs
      const auto& r = from.val.sref.ref;
      create(val.mref.idSheet, r.rwFirst, r.colFirst, r.rwLast, r.colLast);
      break;
    }
    case ExcelType::Ref:
    {
      if (from.val.mref.lpmref->count != 1)
        XLO_THROW("Only contiguous refs supported");
      const auto& r = *from.val.mref.lpmref[0].reftbl;
      create(from.val.mref.idSheet, r.rwFirst, r.colFirst, r.rwLast, r.colLast);
      break;
    }
    default:
      XLO_THROW("ExcelRange: expecting reference type");
    }
  }
  
  inline ExcelRange rangeFromAddress(const wchar_t* address)
  {
    // xlfIndirect will/may core Excel if called outside XLL context!
    if (InXllContext::check())
      return callExcel(msxll::xlfIndirect, address);
    else
      return COM::rangeFromAddress(address);
  }

  XLOIL_EXPORT ExcelRange::ExcelRange(const wchar_t* address)
    : ExcelRange(rangeFromAddress(address))
  {}

  XLOIL_EXPORT ExcelRange::ExcelRange(
    msxll::IDSHEET sheetId, int fromRow, int fromCol, int toRow, int toCol)
  {
    create(sheetId, fromRow, fromCol, toRow, toCol);
  }

  XLOIL_EXPORT std::wstring ExcelRange::address(bool local) const
  {
    auto& ref = this->val.mref.lpmref->reftbl[0];
    if (local)
      return captureStringBuffer([ref](auto buf, auto sz)
        {
          return xlrefToStringA1(ref, buf, sz);
        },
        CELL_ADDRESS_A1_MAX_LEN);
    else
      return captureStringBuffer([this, ref](auto buf, auto sz)
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

  ExcelObj ExcelRange::value() const
  {
    // TODO: does this work outside xll context?
    ExcelObj result;
    callExcelRaw(msxll::xlCoerce, &result, this);
    return result;
  }

  ExcelRange& ExcelRange::operator=(const ExcelObj& value)
  {
    if (InXllContext::check())
    {
      const ExcelObj* args[2];
      args[0] = this;
      args[1] = &value;
      callExcelRaw(msxll::xlSet, nullptr, 2, args);
    }
    else
      COM::rangeSetValue(*this, value);
    return *this;
  }

  void ExcelRange::clear()
  {
    if (InXllContext::check())
      callExcelRaw(msxll::xlSet, nullptr, this);
    else
      COM::rangeSetValue(*this, ExcelObj());
  }
}