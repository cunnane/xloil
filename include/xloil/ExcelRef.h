#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelRange.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Caller.h>

namespace xloil
{
  namespace detail
  {
    template<class TSuper>
    class ExcelRefFn
    {
    public:
      using row_t = Range::row_t;
      using col_t = Range::col_t;

      ExcelObj operator()(int i, int j) const
      {
        return up().range(i, j, i + 1, j + 1).value();
      }

      TSuper& operator=(const ExcelObj& value)
      {
        set(value);
        return up();
      }

      row_t nRows() const
      {
        auto& r = up().ref();
        return r.rwLast - r.rwFirst;
      }
      col_t nCols() const
      {
        auto& r = up().ref();
        return (col_t)(r.colLast - r.colFirst);
      }

      std::tuple<row_t, col_t, row_t, col_t> bounds() const
      {
        auto& r = up().ref();
        return { r.rwFirst, r.colFirst, r.rwLast, r.colLast };
      }

      /// <summary>
      /// Returns the address of the range in the form
      /// 'SheetNm!A1:Z5'
      /// </summary>
      std::wstring address(bool local = false) const
      {
        return local
          ? xlrefLocalAddress(up().ref())
          : xlrefSheetAddress(up().sheetId(), up().ref());
      }

      /// <summary>
      /// Converts the referenced range to an ExcelObj. 
      /// References to single cells return an ExcelObj of the
      /// appropriate type. Multicell refernces return an array.
      /// </summary>
      ExcelObj value() const
      {
        ExcelObj result;
        callExcelRaw(msxll::xlCoerce, &result, &up().obj());
        return result;
      }

      void set(const ExcelObj& value)
      {
        const ExcelObj* args[2];
        args[0] = &up().obj();
        args[1] = &value;
        callExcelRaw(msxll::xlSet, nullptr, 2, args);
      }

      void clear()
      {
        callExcelRaw(msxll::xlSet, nullptr, &up().obj());
      }

    private:
      TSuper& up()             { return (TSuper&)(*this); }
      const TSuper& up() const { return (const TSuper&)(*this); }
    };
  }

  /// <summary>
  /// Contains argument passed to a user-defined function which may be an
  /// ref or sref (local ref) argument. Using this class instead of ExcelObj
  /// in the declaration of a registered function tells xlOil to allow range
  /// references to be passed, otherwise they are converted to arrays.
  /// </summary>
  class RangeArg : public ExcelObj, public detail::ExcelRefFn<RangeArg>
  {
  public:
    friend class detail::ExcelRefFn<RangeArg>;

    RangeArg(const msxll::xlref12& ref)
      : ExcelObj(ref)
    {}

    RangeArg(msxll::IDSHEET sheet, const msxll::xlref12& ref)
      : ExcelObj(sheet, ref)
    {}

    RangeArg range( // TODO: rangearg?
      int fromRow, int fromCol,
      int toRow = Range::TO_END, int toCol = Range::TO_END) const
    {
      switch (xtype())
      {
      case msxll::xltypeRef:
      {
        auto& r = val.mref.lpmref->reftbl[0];
        return RangeArg(val.mref.idSheet, msxll::xlref12{
          r.rwFirst + fromRow,
          r.colFirst + fromCol,
          toRow == Range::TO_END ? r.rwLast : r.rwFirst + toRow,
          toCol == Range::TO_END ? r.colLast : r.colFirst + toCol });
      }
      case msxll::xltypeSRef:
      {
        auto& r = val.sref.ref;
        return RangeArg(msxll::xlref12{
          r.rwFirst + fromRow,
          r.colFirst + fromCol,
          toRow == Range::TO_END ? r.rwLast : r.rwFirst + toRow,
          toCol == Range::TO_END ? r.colLast : r.colFirst + toCol });
      }
      default:
        XLO_THROW("Not a ref");
      }
    }

  protected:
    msxll::IDSHEET sheetId() const
    {
      switch (xtype())
      {
      case msxll::xltypeRef:
        return val.mref.idSheet;
      case msxll::xltypeSRef:
      {
        ExcelObj id;
        callExcelRaw(msxll::xlSheetId, &id);
        return id.val.mref.idSheet;
      }
      default:
        XLO_THROW("Not a ref");
      }
    }

    const msxll::XLREF12& ref() const
    {
      switch (xtype())
      {
      case msxll::xltypeRef:
        return val.mref.lpmref->reftbl[0];
      case msxll::xltypeSRef:
        return val.sref.ref;
      default:
        XLO_THROW("Not a ref");
      }
    }
    msxll::XLREF12& ref()
    {
      switch (xtype())
      {
      case msxll::xltypeRef:
        return val.mref.lpmref->reftbl[0];
      case msxll::xltypeSRef:
        return val.sref.ref;
      default:
        XLO_THROW("Not a ref");
      }
    }

    const ExcelObj& obj() const { return *this; }
    ExcelObj&       obj()       { return *this; }
  };


  /// <summary>
  /// Normalises a reference to a range on an Excel sheet, i.e. taken XLL 
  /// ref or sref (local reference, i.e. to the active sheet) type ExcelObj 
  /// and turns it into a global reference
  /// </summary>
  class XLOIL_EXPORT ExcelRef : public detail::ExcelRefFn<ExcelRef>
  {
  public:
    using row_t = int;
    using col_t = int;
    friend class detail::ExcelRefFn<ExcelRef>;

    /// <summary>
    /// Constructs an ExcelRange from an ExcelObj. Will throw if
    /// the object is not of type Ref or SRef.
    /// </summary>
   ExcelRef(const ExcelObj& from);

    /// <summary>
    /// Constructs an ExcelRange from a sheet address. If the 
    /// address does not contain a sheet name, the current
    /// Active sheet is used.
    /// </summary>
    explicit ExcelRef(const wchar_t* address);

    ExcelRef(msxll::IDSHEET sheetId,
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

    ExcelRef range(
      int fromRow, int fromCol,
      int toRow = Range::TO_END, int toCol = Range::TO_END) const
    {
      return ExcelRef(sheetId(),
        ref().rwFirst + fromRow,
        ref().colFirst + fromCol,
        toRow == Range::TO_END ? ref().rwLast : ref().rwFirst + toRow,
        toCol == Range::TO_END ? ref().colLast : ref().colFirst + toCol);
    }

    operator const ExcelObj& () const { return _obj; }
    

  protected:
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
    
    const ExcelObj& obj() const { return _obj; }
    ExcelObj&       obj()       { return _obj; }

  private:
    ExcelObj _obj;
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
    explicit XllRange(const ExcelObj& ref);

    virtual Range* range(
      int fromRow, int fromCol,
      int toRow = TO_END, int toCol = TO_END) const final;

    virtual std::tuple<row_t, col_t> shape() const final;

    virtual std::tuple<row_t, col_t, row_t, col_t> bounds() const final;

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