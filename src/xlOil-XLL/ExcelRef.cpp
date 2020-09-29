#include <xloil/ExcelRef.h>


namespace xloil
{
  XLOIL_EXPORT ExcelRef::ExcelRef(const ExcelObj& from)
  {
    switch (from.type())
    {
    case ExcelType::SRef:
    {
      // TODO: this may not work as expected in macro funcs
      if (0 != callExcelRaw(msxll::xlSheetId, &_obj))
        XLO_THROW("ExcelRef: call to xlSheetId failed");
      const auto& r = from.val.sref.ref;
      create(_obj.val.mref.idSheet, r.rwFirst, r.colFirst, r.rwLast, r.colLast);
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

  XLOIL_EXPORT ExcelRef::ExcelRef(const wchar_t* address)
    : ExcelRef(callExcel(msxll::xlfIndirect, address))
  {}

  XLOIL_EXPORT ExcelRef::ExcelRef(
    msxll::IDSHEET sheetId, int fromRow, int fromCol, int toRow, int toCol)
  {
    create(sheetId, fromRow, fromCol, toRow, toCol);
  }
  void ExcelRef::create(
    msxll::IDSHEET sheetId, 
    row_t fromRow, col_t fromCol, 
    row_t toRow, col_t toCol)
  {
    _obj.xltype = msxll::xltypeRef | msxll::xlbitDLLFree;
    _obj.val.mref.idSheet = sheetId;
    _obj.val.mref.lpmref = new msxll::XLMREF12[1];
    _obj.val.mref.lpmref->count = 1;
    ref().rwFirst = fromRow;
    ref().colFirst = fromCol;
    ref().rwLast = toRow;
    ref().colLast = toCol;
    if (toRow >= XL_MAX_ROWS || fromRow > toRow)
      XLO_THROW("ExcelRef out of range fromRow={0}, toRow={1}", fromRow, toRow);
    if (toCol >= XL_MAX_COLS || fromCol > toCol)
      XLO_THROW("ExcelRef out of range fromCol={0}, toCol={1}", fromCol, toCol);
  }
  XllRange::XllRange(const ExcelRef& ref)
    : _ref(ref)
  {}
  Range* XllRange::range(int fromRow, int fromCol, int toRow, int toCol) const
  {
    return new XllRange(_ref.range(fromRow, fromCol, toRow, toCol));
  }

  std::tuple<Range::row_t, Range::col_t> XllRange::shape() const
  {
    return { _ref.nRows(), _ref.nCols() };
  }

  std::tuple<Range::row_t, Range::col_t, Range::row_t, Range::col_t> XllRange::bounds() const
  {
    return _ref.bounds();
  }

  std::wstring XllRange::address(bool local) const
  {
    return _ref.address(local);
  }

  ExcelObj XllRange::value() const
  {
    return _ref.value();
  }

  ExcelObj XllRange::value(row_t i, col_t j) const
  {
    return _ref.range(i, j, i + 1, j + 1).value();
  }

  void XllRange::set(const ExcelObj & value)
  {
    _ref.set(value);
  }

  void XllRange::clear()
  {
    _ref.clear();
  }

  Range* newXllRange(const ExcelObj& xlRef)
  {
    return new XllRange(xlRef);
  }
}