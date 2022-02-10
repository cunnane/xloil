#pragma once
#include <xlOil/ExportMacro.h>
#include <xlOil/XlCallSlim.h>
#include <xlOil/PString.h>
#include <xlOil/ExcelObj.h>
#include <memory>
#include <string>

namespace xloil
{
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
    /// Species the format used to write sheet addresses
    /// </summary>
    enum AddressStyle
    {
      /// <summary>
      /// A1 Format: [Book1]Sheet1!A1:B2
      /// </summary>
      A1,
      /// <summary>
      /// RC Format: [Book1]Sheet1!R1C1:R2C2
      /// </summary>
      RC,
    };

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
    /// sref. If it corresponds to a valid Excel sheet, the sheetId is looked
    /// up and can be used in an Internal-style reference.
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
    /// <param name="buf"></param>
    /// <param name="bufLen"></param>
    /// <param name="style">Selects A1-type, RC-type or internal address</param>
    /// <returns></returns>
    int writeAddress(wchar_t* buf, size_t bufLen, AddressStyle style = RC) const;

    /// <summary>
    /// As per <see cref="writeAddress"/>, but returns a string rather than writing
    /// to a buffer
    /// </summary>
    /// <param name="style"></param>
    /// <returns></returns>
    std::wstring writeAddress(AddressStyle style = RC) const;

    /// <summary>
    /// Returns the calling worksheet name as a PString or a null PString
    /// if it could not be determined.
    /// </summary>
    /// <returns></returns>
    PStringView<> fullSheetName() const
    {
      return _sheetName.asPString();
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
      switch (_address.type())
      {
      case ExcelType::SRef: return &_address.val.sref.ref;
      default:
        return nullptr;
      }
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
  /// to the provided string buffer. That is, gives 'Sheet!A1' or 'Sheet!R1C1'.
  /// <returns>The number of characters written</returns>
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefWriteWorkbookAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    bool A1Style = true);

  /// <summary>
  /// Version of <see cref="xlrefToWorkbookAddress"/> which returns a string rather
  /// than writing to a buffer
  /// </summary>
  inline std::wstring xlrefToWorkbookAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    bool A1Style = true)
  {
    return captureWStringBuffer([&](auto buf, auto sz)
    {
      return xlrefWriteWorkbookAddress(sheet, ref, buf, sz, A1Style);
    });
  }
  
  /// <summary>
  /// Writes a simple Excel ref (not including sheet name) to 'RxCy' or 
  /// 'RaCy:RxCy' format in the provided string buffer. 
  /// <returns>The number of characters written</returns>
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefToLocalRC(
    const msxll::XLREF12& ref, 
    wchar_t* buf,
    size_t bufSize);

  /// <summary>
  /// Writes a local Excel ref (not including sheet name) to 'A1' or 'A1:Z9' 
  /// format in the provided string buffer. Writes a null-terminator.
  /// <returns>
  ///   The number of characters written not including the null terminator or zero
  ///   if the bufSize is insufficient.
  /// </returns>
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefToLocalA1(
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize);

  /// <summary>
  /// Similar to <see cref="xlrefToWorkbookAddress"/>, but without the sheet name
  /// </summary>
  inline std::wstring xlrefToLocalAddress(
    const msxll::XLREF12& ref,
    bool A1Style = true)
  {
    return captureWStringBuffer([&](auto buf, auto sz)
      {
        return A1Style
          ? xlrefToLocalA1(ref, buf, sz)
          : xlrefToLocalRC(ref, buf, sz);
      },
      XL_CELL_ADDRESS_A1_MAX_LEN);
  }

  /// <summary>
  /// Parses a local Excel ref (not including sheet name) such as 'A1' or 'A1:Z9'
  /// to an XLREF12 object. Returns false if the string could not be parsed into
  /// a valid XLREF12 and sets the offending members to -1.
  /// </summary>
  /// <param name="r"></param>
  /// <param name="address"></param>
  /// <returns></returns>
  XLOIL_EXPORT bool localAddressToXlRef(
    msxll::XLREF12& r,
    const std::wstring_view& address);

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