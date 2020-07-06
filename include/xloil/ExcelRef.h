#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelRange.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Caller.h>

namespace xloil
{
  /// <summary>
  /// Wraps a reference to a range on an Excel sheet, i.e. an XLL 
  /// ref or sref (local reference) type ExcelObj.
  /// </summary>
  class ExcelRef
  {
  public:
    using row_t = Range::row_t;
    using col_t = Range::col_t;

    /// <summary>
    /// Constructs an ExcelRange from an ExcelObj. Will throw if
    /// the object is not of type Ref or SRef.
    /// </summary>
    XLOIL_EXPORT ExcelRef(const ExcelObj& from);

    /// <summary>
    /// Constructs an ExcelRange from a sheet address. If the 
    /// address does not contain a sheet name, the current
    /// Active sheet is used.
    /// </summary>
    XLOIL_EXPORT explicit ExcelRef(const wchar_t* address);

    XLOIL_EXPORT ExcelRef(msxll::IDSHEET sheetId,
      int fromRow, int fromCol, int toRow, int toCol);

    /// <summary>
    /// Copy constructor
    /// </summary>
    ExcelRef(const ExcelRef& from)
      : _obj(from._obj)
    {}

    ~ExcelRef()
    {
      reset();
    }

    ExcelRef& operator=(const ExcelObj& value)
    {
      set(value);
    }

    ExcelObj operator()(int i, int j) const
    {
      return range(i, j, i+1, j+1).value();
    }

    ExcelRef range(
      int fromRow, int fromCol,
      int toRow = Range::TO_END, int toCol = Range::TO_END) const
    {
      // Excel's ranges are _inclusive_ at the right hand end. This is
      // unusual in programming languages, so we hide it by adjusting
      // toRow / toCol here
      return ExcelRef(sheetId(),
        ref().rwFirst + fromRow,
        ref().colFirst + fromCol,
        toRow < 0 ? ref().rwLast + toRow + 1 : ref().rwFirst + toRow - 1,
        toCol < 0 ? ref().colLast + toCol + 1 : ref().colFirst + toCol - 1);
    }

    row_t nRows() const
    {
      return ref().rwLast - ref().rwFirst;
    }
    col_t nCols() const
    {
      return (col_t)(ref().colLast - ref().colFirst);
    }

    /// <summary>
    /// Returns the address of the range in the form
    /// 'SheetNm!A1:Z5'
    /// </summary>
    std::wstring address(bool local = false) const
    {
      return local 
        ? xlrefLocalAddress(ref()) 
        : xlrefSheetAddress(sheetId(), ref());
    }

    /// <summary>
    /// Converts the referenced range to an ExcelObj. 
    /// References to single cells return an ExcelObj of the
    /// appropriate type. Multicell refernces return an array.
    /// </summary>
    ExcelObj ExcelRef::value() const
    {
      ExcelObj result;
      callExcelRaw(msxll::xlCoerce, &result, &_obj);
      return result;
    }

    ExcelRef& ExcelRef::set(const ExcelObj& value)
    {
      const ExcelObj* args[2];
      args[0] = &_obj;
      args[1] = &value;
      callExcelRaw(msxll::xlSet, nullptr, 2, args);
      return *this;
    }

    void ExcelRef::clear()
    {
      callExcelRaw(msxll::xlSet, nullptr, &_obj);
    }

    const ExcelObj& asExcelObj() const { return _obj; }

  private:
    ExcelObj _obj;

    msxll::IDSHEET  sheetId() const { return _obj.val.mref.idSheet; }
    msxll::IDSHEET& sheetId()       { return _obj.val.mref.idSheet; }

    const msxll::XLREF12& ref() const
    {
      return _obj.val.mref.lpmref->reftbl[0];
    }

    msxll::XLREF12& ref()
    {
      return _obj.val.mref.lpmref->reftbl[0];
    }

    void create(
      msxll::IDSHEET sheetId, 
      row_t fromRow, col_t fromCol,
      row_t toRow, col_t toCol);

    void reset()
    {
      if (_obj.xltype & msxll::xlbitDLLFree)
      {
        delete[] _obj.val.mref.lpmref;
        _obj.xltype = msxll::xltypeNil;
      }
    }
  };

  /// <summary>
  /// An implementation of Range which uses an ExcelRef, i.e. an Xll sheet 
  /// reference as it's underlying type
  /// </summary>
  class XllRange : public Range
  {
  public:
    explicit XllRange(const ExcelRef& ref);

    virtual Range* range(
      int fromRow, int fromCol,
      int toRow = TO_END, int toCol = TO_END) const final;

    virtual row_t nRows() const final;

    virtual col_t nCols() const final;

    /// <summary>
    /// Returns the address of the range in the form
    /// 'SheetNm!A1:Z5'
    /// </summary>
    virtual std::wstring address(bool local = false) const final;

    /// <summary>
    /// Converts the referenced range to an ExcelObj. 
    /// References to single cells return an ExcelObj of the
    /// appropriate type. Multicell refernces return an array.
    /// </summary>
    virtual ExcelObj value() const final;

    virtual ExcelObj value(row_t i, col_t j) const final;

    virtual void set(const ExcelObj& value) final;

    /// <summary>
    /// Clears / empties all cells referred to by this ExcelRange.
    /// </summary>
    virtual void clear() final;

    const ExcelRef& native() const { return _ref; }

  private:
    ExcelRef _ref;
  };
}