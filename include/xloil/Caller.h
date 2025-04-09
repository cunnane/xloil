#pragma once
#include <xlOil/ExportMacro.h>
#include <xlOil/XlCallSlim.h>
#include <xlOil/PString.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/EnumHelper.h>
#include <memory>
#include <string>

namespace xloil
{
  /// <summary>
   /// Species the format used to write sheet addresses
   /// </summary>
  enum class AddressStyle : int
  {
    /// <summary>
    /// A1 Format: '[Book1]Sheet1'!A1:B2
    /// </summary>
    A1 = 0,
    /// <summary>
    /// RC Format: '[Book1]Sheet1'!R1C1:R2C2
    /// </summary>
    RC = 1,
    /// <summary>
    /// Makes the address absolute, e.g. $A$1
    /// </summary>
    ABSOLUTE = 2,
    /// <summary>
    /// Does not quote sheet name, e.g. [Book1]Sheet1!A1:B2
    /// </summary>
    NOQUOTE = 4,
  };

  /// <summary>
  /// Captures and writes information about the calling cell or context 
  /// provided by xlfCaller. Only returns useful information when the
  /// caller was a worksheet
  /// </summary>
  class XLOIL_EXPORT CallerInfo
  {
  private:    
    ExcelObj _address;
    ExcelObj _sheetName;

  public:
    /// <summary>
    /// Constructor which makes calls to xlfCaller and xlfSheetName to
    /// determine the caller.
    /// </summary>
    CallerInfo();

    /// <summary>
    /// Provide custom caller information. The <paramref name="address"/> is
    /// interpreted as per the return from xlfCaller. In particular, a string
    /// address will be returned by <see cref="writeAddress"/> unmodified. The 
    /// <paramref name="fullSheetName"/> is used when the address is a ref or
    /// sref.
    /// </summary>
    /// <param name="address"></param>
    /// <param name="fullSheetName">If provided, should be of the form [Book]Sheet</param>
    CallerInfo(const ExcelObj& address, const wchar_t* fullSheetName=nullptr);

    /// <summary>
    /// Returns the upper bound on the string length required to write the address
    /// </summary>
    /// <returns></returns>
    uint16_t addressLength(AddressStyle style) const;

    /// <summary>
    /// Writes the caller address to the provided buffer, returning the number
    /// of characters written on success or a negative number or on failure. 
    /// </summary>
    /// <param name="style">Selects A1-type or RC-type</param>
    /// <returns></returns>
    int writeAddress(
      wchar_t* buf, 
      size_t bufLen, 
      AddressStyle style = AddressStyle::A1) const;

    /// <summary>
    /// As per <see cref="writeAddress"/>, but returns a string rather than writing
    /// to a buffer
    /// </summary>
    /// <param name="style"></param>
    /// <returns></returns>
    std::wstring address(AddressStyle style = AddressStyle::A1) const;


    std::wstring localAddress(AddressStyle style = AddressStyle::A1) const;

    /// <summary>
    /// Returns the calling worksheet name as a PString or a null PString
    /// if it could not be determined.
    /// </summary>
    /// <returns></returns>
    PStringRef fullSheetName() const
    {
      return _sheetName.cast<PStringRef>();
    }
   
    /// <summary>
    /// Returns a view containing only the sheet name.
    /// </summary>
    std::wstring_view sheetName() const
    {
      auto sName = fullSheetName();
      if (sName.empty())
        return std::wstring_view();
      auto begin = sName.begin() + sName.find(L']') + 1;
      return std::wstring_view(begin, sName.end() - begin);
    }

    /// <summary>
    /// Returns a view containing only the workbook name. If the workbook has
    /// been saved, this includes a file extension.
    /// </summary>
    std::wstring_view workbook() const
    {
      auto sName = fullSheetName();
      if (sName.empty())
        return std::wstring_view();
      return std::wstring_view(sName.begin() + 1, sName.rfind(L']') - 1);
    }

    /// <summary>
    /// Returns a pointer to a XLREF12 sheet reference if caller was a 
    /// worksheet, else returns nullptr.
    /// </summary>
    const msxll::XLREF12* sheetRef() const
    {
      return _address.isType(ExcelType::SRef)
        ? &_address.val.sref.ref
        : nullptr;
    }
  };

  /// <summary>
  /// Returns the Excel A1-style column letter corresponding
  /// to a given zero-based column index. Returns the number of
  /// characters written (1, 2 or 3)
  /// </summary>
  XLOIL_EXPORT uint8_t writeColumnName(size_t colIndex, char buf[3]);

  /// <summary>
  /// Writes a simple Excel ref including sheet name in either A1 or RxCy 
  /// to the provided string buffer. That is, gives "[Book]Sheet!A1" or 
  /// "[Book]Sheet!R1C1".  The workbook and sheet name may be quoted which gives
  /// some consistency with COM's Range.Address, however in the COM case 
  /// there is a complex set of rules determining whether quotes are added.
  /// <returns>The number of characters written</returns>
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefWriteWorkbookAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    AddressStyle style = AddressStyle::A1);

  /// <summary>
  /// Version of <see cref="xlrefToWorkbookAddress"/> which returns a string rather
  /// than writing to a buffer
  /// </summary>
  inline std::wstring xlrefToWorkbookAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    AddressStyle style = AddressStyle::A1)
  {
    return captureWStringBuffer([&](auto buf, auto sz)
    {
      return xlrefWriteWorkbookAddress(sheet, ref, buf, sz, style);
    });
  }


  /// <summary>
  /// Writes an Excel ref to 'A1'/'A1:Z9' or 'RxCy'/'RaCy:RxCy' format
  /// in the provided string buffer. Includes a null-terminator.
  /// <returns>
  ///   The number of characters written not including the null terminator or 
  ///   zero if the bufSize is insufficient.
  /// </returns>
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefToAddress(
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    const std::wstring_view& sheetName = std::wstring_view(),
    AddressStyle style = AddressStyle::A1);

  /// <summary>
  /// Version of <see cref="xlrefToAddress"/> which writes to a string
  /// </summary>
  inline std::wstring xlrefToAddress(
    const msxll::XLREF12& ref,
    const std::wstring_view& sheetName = std::wstring_view(),
    AddressStyle style = AddressStyle::A1)
  {
    return captureWStringBuffer([&](auto buf, auto sz)
      {
        return xlrefToAddress(ref, buf, sz, sheetName, style);
      },
      XL_CELL_ADDRESS_A1_MAX_LEN);
  }

  /// <summary>
  /// Parses a local Excel address (not including sheet name) from a string such as 
  /// 'A1' or 'A1:Z9' or 'R1C4' to an XLREF12 object. Returns false if the string 
  /// could not be parsed into a valid XLREF12.
  /// </summary>
  bool localAddressToXlRef(
    const std::wstring_view& address,
    msxll::XLREF12& result);

  /// <summary>
  /// Parses an Excel address including sheet name from a string such as 
  /// 'A1' or 'A1:Z9' or 'R1C4' to an XLREF12 object. Returns false if the string 
  /// could not be parsed into a valid XLREF12.
  /// </summary>
  XLOIL_EXPORT bool addressToXlRef(
    const std::wstring_view& address,
    msxll::XLREF12& result,
    std::wstring* sheetName = nullptr);

  /// <summary>
  /// Returns true if the user is currently in the function wizard.
  /// It's quite an expensive check involving looping through visible 
  /// Windows as Excel does not provide a built-in way to check this.
  /// </summary>
  XLOIL_EXPORT bool inFunctionWizard();

  /// <summary>
  /// Throws "#WIZARD!" true if the user is currently in the function 
  /// wizard.  The idea being that this string will be returned to Excel
  /// by the surrounding try...catch.
  /// 
  /// Quite an expensive check as Excel does not provide a built-in 
  /// way to check this.
  /// </summary>
  inline void throwInFunctionWizard()
  {
    if (xloil::inFunctionWizard())
      throw std::runtime_error("#WIZARD!");
  }
}