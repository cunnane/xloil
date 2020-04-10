#pragma once
#include "ExcelObj.h"
#include "ExcelCall.h"
#include "ExcelState.h"
#include "ExportMacro.h"

namespace xloil
{
  class ExcelRange : protected ExcelObj
  {
  public:
    using row_t = uint32_t;
    using col_t = uint16_t;

    XLOIL_EXPORT ExcelRange(const ExcelObj& from);

    XLOIL_EXPORT ExcelRange(const wchar_t* address);

    XLOIL_EXPORT ExcelRange(msxll::IDSHEET sheetId, 
      int fromRow, int fromCol, int toRow, int toCol);

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
      return ref().colLast - ref().colFirst;
    }
    size_t size() const
    {
      return nRows() * nCols();
    }

    ExcelObj value() const
    {
      ExcelObj result;
      callExcelRaw(msxll::xlCoerce, &result, this);
      return result;
    }

    XLOIL_EXPORT std::wstring address(bool local = false) const;
 
    ExcelRange& operator=(const ExcelObj& value)
    {
      const ExcelObj* args[2];
      args[0] = this;
      args[1] = &value;
      callExcelRaw(msxll::xlSet, nullptr, 2, args);
      return *this;
    }

    void clear()
    {
      callExcelRaw(msxll::xlSet, nullptr, this);
    }

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
