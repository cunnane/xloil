#include <xloil/ExcelRef.h>


namespace xloil
{
  XLOIL_EXPORT ExcelRef::ExcelRef(const ExcelObj& from)
  {
    switch (from.type())
    {
    case ExcelType::SRef:
    {
      // If we don't know th sheet ID, we first call xlSheetNm to get 
      // the 'current' i.e. currently calculating sheet, then fetch 
      // the sheet ID.  Calling xlSheetId directly with no argument 
      // gets the front sheet, i.e. topmost window or 'active' sheet.
      ExcelObj sheetName;
      callExcelRaw(msxll::xlSheetNm, &sheetName, &from);
      if (0 != callExcelRaw(msxll::xlSheetId, &_obj, &sheetName))
        XLO_THROW("ExcelRef: could not determine sheet for local reference");
      create(_obj.val.mref.idSheet, from.val.sref.ref);
      break;
    }
    case ExcelType::Ref:
    {
      if (from.val.mref.lpmref->count != 1)
        XLO_THROW("ExcelRef: only contiguous refs are supported");
      create(from.val.mref.idSheet, *from.val.mref.lpmref[0].reftbl);
      break;
    }
    default:
      XLO_THROW("ExcelRange: expecting reference type");
    }
  }

  XLOIL_EXPORT ExcelRef::ExcelRef(const std::wstring_view& address)
  {
    // If address contains a '!', get sheetId of that name otherwise use
    // the active sheet (which we get by passing no args). Sheet name
    // may be quoted - xlSheetId doesn't like this so we must de-quote
    const auto pling = address.find_last_of(L'!');
    const auto quoted = address[0] == L'\'' ? 1 : 0;
    auto [sheetId, ret] = pling > 0
      ? tryCallExcel(msxll::xlSheetId, address.substr(0 + quoted, pling - quoted * 2))
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
    msxll::IDSHEET sheetId, const msxll::xlref12& ref)
  {
    create(sheetId, ref);
  }

  void ExcelRef::create(
    msxll::IDSHEET sheetId, 
    const msxll::xlref12& ref)
  {
    _obj = ExcelObj(sheetId, ref);
  }

  void XllRange::setFormula(const std::wstring_view& formula, bool array)
  {
    // Formulae must use RC style references
    if (size() > 1 && array)
      callExcel(msxll::xlcFormulaArray, formula, _ref);
    else
      callExcel(msxll::xlcFormula, formula, _ref);
  }
}