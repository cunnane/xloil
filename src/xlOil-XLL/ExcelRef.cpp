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

  XLOIL_EXPORT ExcelRef::ExcelRef(const std::wstring_view& address)
    : ExcelRef(callExcel(msxll::xlfIndirect, address))
  {
    // If address contains a '!', get sheetId of that name otherwise use
    // the active sheet (which we get by passing no args)
    auto pling = address.find_last_of(L'!');
    auto [sheetId, ret] = pling > 0
      ? tryCallExcel(msxll::xlSheetId, address.substr(0, pling))
      : tryCallExcel(msxll::xlSheetId);

    if (ret != 0 || !sheetId.isType(ExcelType::Ref))
      XLO_THROW(L"Could not find sheet name from address {}", address);

    msxll::XLREF12 sref;
    if (!localAddressToXlRef(sref, address.substr(pling + 1)))
    {
      // If the address cannot be parsed, it may be RC format or an Excel
      // range name, so we call the API to resolve it. Note, indirect only  
      // returns an sref, even if the address contains a sheet name.
      sref = callExcel(msxll::xlfIndirect, address).val.sref.ref;
    }
  
    _obj = ExcelObj(sheetId.val.mref.idSheet, sref);
  }

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
    _obj = ExcelObj(sheetId, msxll::xlref12{ fromRow,  toRow, fromCol, toCol });
  }

  void XllRange::setFormula(const std::wstring_view& formula)
  {
    // Formulae must use RC style references
    if (size() > 1)
      callExcel(msxll::xlcFormulaArray, formula, _ref);
    else
      callExcel(msxll::xlcFormula, formula, _ref);
  }
}