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

    static constexpr int TO_END = 0;

    // Doesn't check that a sub-range has been specified
    ExcelRange range(int fromRow, int fromCol, int toRow = TO_END, int toCol = TO_END) const
    {
      return ExcelRange(sheetId(),
        ref().rwFirst + fromRow, 
        ref().colFirst + fromCol,
        toRow <= 0 ? ref().rwLast + toRow : ref().rwFirst + toRow,
        toCol <= 0 ? ref().colLast + toCol : ref().colFirst + toCol);
    }
    ExcelRange cell(int i, int j)
    {
      return range(i, j, 1, 1);
    }

    size_t nRows() const 
    {
      return ref().rwLast - ref().rwFirst;
    }
    size_t nCols() const 
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
