#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/ExportMacro.h>

namespace xloil
{
  /// <summary>
  /// An ExcelRange holds an Excel sheet reference and provides
  /// functionality to access it. ExcelRange can only be used by
  /// macro-enabled functions.
  /// </summary>
  class ExcelRange : protected ExcelObj
  {
  public:
    using row_t = uint32_t;
    using col_t = uint16_t;

    /// <summary>
    /// Constructs an ExcelRange from an ExcelObj. Will throw if
    /// the object is not of type Ref or SRef.
    /// </summary>
    XLOIL_EXPORT ExcelRange(const ExcelObj& from);

    /// <summary>
    /// Constructs an ExcelRange from a sheet address. If the 
    /// address does not contain a sheet name, the current
    /// Active sheet is used.
    /// </summary>
    XLOIL_EXPORT ExcelRange(const wchar_t* address);

    XLOIL_EXPORT ExcelRange(msxll::IDSHEET sheetId, 
      int fromRow, int fromCol, int toRow, int toCol);

    /// <summary>
    /// Copy constructor
    /// </summary>
    ExcelRange(const ExcelRange& from)
      : ExcelObj(static_cast<const ExcelObj&>(from))
    {}

    ~ExcelRange()
    {
      reset();
    }

    ExcelObj operator()(int i, int j) const
    {
      return ExcelRange(sheetId(), i, j, i + 1, j + 1).value();
    }

    static constexpr int TO_END = -1;

    /// <summary>
    /// Gives a subrange relative to the current range. Unlike Excel's VBA Range function
    /// we used zero-based indexing and do not include the right-hand endpoint.
    /// Similar to Excel's function, we do not insist the sub-range is a subset, so
    /// fromRow can be negative or toRow can be past the end of the referenced range.
    /// </summary>
    /// <param name="fromRow"></param>
    /// <param name="fromCol"></param>
    /// <param name="toRow"></param>
    /// <param name="toCol"></param>
    /// <returns></returns>
    ExcelRange range(int fromRow, int fromCol, int toRow = TO_END, int toCol = TO_END) const
    {
      // Excel's ranges are _inclusive_ at the right hand end. This 
      // is unusual in programming languages, so we hide it by 
      // adjusting toRow / toCol here
      return ExcelRange(sheetId(),
        ref().rwFirst + fromRow, 
        ref().colFirst + fromCol,
        toRow < 0 ? ref().rwLast + toRow + 1 : ref().rwFirst + toRow - 1,
        toCol < 0 ? ref().colLast + toCol + 1 : ref().colFirst + toCol - 1);
    }
    /// <summary>
    /// Returns a 1x1 subrange containing the specified cell. Uses zero-based
    /// indexing unlike Excel's VBA Range.Cells function.
    /// </summary>
    /// <param name="i"></param>
    /// <param name="j"></param>
    /// <returns></returns>
    ExcelRange cells(int i, int j) const
    {
      return range(i, j, i + 1, j + 1);
    }

    row_t nRows() const
    {
      return ref().rwLast - ref().rwFirst;
    }
    col_t nCols() const 
    {
      return (col_t)(ref().colLast - ref().colFirst);
    }
    size_t size() const
    {
      return nRows() * nCols();
    }

    /// <summary>
    /// Returns the address of the range in the form
    /// 'SheetNm!A1:Z5'
    /// </summary>
    XLOIL_EXPORT std::wstring address(bool local = false) const;

    /// <summary>
    /// Converts the referenced range to an ExcelObj. 
    /// References to single cells return an ExcelObj of the
    /// appropriate type. Multicell refernces return an array.
    /// </summary>
    XLOIL_EXPORT ExcelObj value() const;

    XLOIL_EXPORT ExcelRange& operator=(const ExcelObj& value);

    /// <summary>
    /// Clears / empties all cells referred to by this ExcelRange.
    /// </summary>
    XLOIL_EXPORT void clear();

    msxll::IDSHEET sheetId() const 
    {
      return val.mref.idSheet;
    }

  private:
    const msxll::XLREF12& ref() const
    {
      return val.mref.lpmref->reftbl[0];
    }
    msxll::XLREF12& ref()
    {
      return val.mref.lpmref->reftbl[0];
    }
  
    msxll::IDSHEET& sheetId() 
    {
      return val.mref.idSheet;
    }
    
    void create(
      msxll::IDSHEET sheetId, int fromRow, int fromCol, int toRow, int toCol);

    void reset()
    {
      if (xltype & msxll::xlbitDLLFree)
      {
        delete[] val.mref.lpmref;
        xltype = msxll::xltypeNil;
      }
    }
  };
}
