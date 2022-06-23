#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Range.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Caller.h>
#include <xlOil/ExcelArray.h>

namespace xloil
{
  namespace detail
  {
    template<class TSuper>
    class ExcelRefArgBase
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

      std::tuple<row_t, col_t> shape() const
      {
        return { nRows(), nCols() };
      }

      std::tuple<row_t, col_t, row_t, col_t> bounds() const
      {
        auto& r = up().ref();
        return { r.rwFirst, r.colFirst, r.rwLast, r.colLast };
      }

      /// <summary>
      /// Returns the address of the range in the form '[Book]SheetNm'!A1:Z5
      /// </summary>
      std::wstring address(bool local = false) const
      {
        return local
          ? xlrefToLocalAddress(up().ref())
          : xlrefToWorkbookAddress(up().sheetId(), up().ref());
      }

      /// <summary>
      /// Converts the referenced range to an ExcelObj. References to single
      /// cells return an ExcelObj of the appropriate type. Multicell refernces 
      /// return an array.
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

      static auto subrangeHelper(
        const msxll::xlref12& r,
        int fromRow, int fromCol,
        int toRow, int toCol) noexcept
      {
        return msxll::xlref12{
          r.rwFirst + fromRow,
          toRow == Range::TO_END ? r.rwLast : r.rwFirst + toRow,
          r.colFirst + fromCol,
          toCol == Range::TO_END ? r.colLast : r.colFirst + toCol };
      }

    private:
      TSuper& up()             noexcept { return (TSuper&)(*this); }
      const TSuper& up() const noexcept { return (const TSuper&)(*this); }
    };
  }


  /// <summary>
  /// Normalises a reference to a range on an Excel sheet, i.e. taken XLL 
  /// ref or sref (local reference, i.e. to the active sheet) type ExcelObj 
  /// and turns it into a global reference
  /// </summary>
  class XLOIL_EXPORT ExcelRef : public detail::ExcelRefArgBase<ExcelRef>
  {
  public:
    using row_t = int;
    using col_t = int;
    friend class detail::ExcelRefArgBase<ExcelRef>;

    /// <summary>
    /// Constructs an ExcelRange from an ExcelObj. Will throw if
    /// the object is not of type Ref or SRef.
    /// </summary>
    ExcelRef(const ExcelObj& from);

    /// <summary>
    /// Constructs an ExcelRange from a sheet address. If the address
    /// does not contain a sheet name, the current Active sheet is used.
    /// </summary>
    explicit ExcelRef(const std::wstring_view& address);
    explicit ExcelRef(const wchar_t* address) 
      : ExcelRef(std::wstring_view(address)) 
    {}

    ExcelRef(msxll::IDSHEET sheetId, const msxll::xlref12& ref);

    ExcelRef(msxll::IDSHEET sheetId,
      int fromRow, int fromCol,
      int toRow, int toCol)
      : ExcelRef(sheetId, msxll::xlref12{ fromRow, toRow, fromCol, toCol })
    {}

    /// <summary>
    /// Copy constructor
    /// </summary>
    ExcelRef(const ExcelRef& from)
      : _obj(from._obj)
    {}

    ExcelRef(ExcelRef&& from) noexcept
      : _obj(std::move(from._obj))
    {}

    ~ExcelRef() noexcept
    {
      reset();
    }

    ExcelRef range(
      int fromRow, int fromCol,
      int toRow = Range::TO_END, int toCol = Range::TO_END) const
    {
      return ExcelRef(sheetId(), 
        subrangeHelper(ref(), fromRow, fromCol, toRow, toCol));
    }

    operator const ExcelObj& () const noexcept { return _obj; }


  protected:
    msxll::IDSHEET  sheetId() const noexcept { return _obj.val.mref.idSheet; }
    msxll::IDSHEET& sheetId()       noexcept { return _obj.val.mref.idSheet; }

    const msxll::XLREF12& ref() const noexcept
    {
      return _obj.val.mref.lpmref->reftbl[0];
    }
    msxll::XLREF12& ref() noexcept
    {
      return _obj.val.mref.lpmref->reftbl[0];
    }

    const ExcelObj& obj() const noexcept { return _obj; }
    ExcelObj&       obj()       noexcept { return _obj; }

  private:
    ExcelObj _obj;

    void create(
      msxll::IDSHEET sheetId,
      const msxll::xlref12& ref);

    void reset() noexcept
    {
      if (_obj.xltype & msxll::xlbitDLLFree)
      {
        delete[] _obj.val.mref.lpmref;
        _obj.xltype = msxll::xltypeNil;
      }
    }
  };


  /// <summary>
  /// Contains argument passed to a user-defined function which may be an
  /// ref or sref (local ref) argument. Using this class instead of ExcelObj
  /// in the declaration of a registered function tells xlOil to allow range
  /// references to be passed, otherwise they are converted to arrays.
  /// </summary>
  class RangeArg : public ExcelObj, public detail::ExcelRefArgBase<RangeArg>
  {
    friend class detail::ExcelRefArgBase<RangeArg>;

    /// <summary>
    /// Not externally constructable. Prefer to store or pass a ExcelRef: 
    /// this avoids inadvertent use of the local range (SRef) type which doesn't
    /// link to a specific sheet.
    /// </summary>
    /// <param name="ref"></param>
    RangeArg(const msxll::xlref12& ref)
      : ExcelObj(ref)
    {}

    RangeArg(msxll::IDSHEET sheet, const msxll::xlref12& ref)
      : ExcelObj(sheet, ref)
    {}

  public:
    ExcelRef toExcelRef()
    {
      return ExcelRef(*this);
    }

    RangeArg range(
      int fromRow, int fromCol,
      int toRow = Range::TO_END, int toCol = Range::TO_END) const
    {
      switch (xtype())
      {
      case msxll::xltypeRef:
      {
        auto& r = val.mref.lpmref->reftbl[0];
        return RangeArg(val.mref.idSheet, 
          subrangeHelper(r, fromRow, fromCol, toRow, toCol));
      }
      case msxll::xltypeSRef:
      {
        auto& r = val.sref.ref;
        return RangeArg(subrangeHelper(r, fromRow, fromCol, toRow, toCol));
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

    const ExcelObj& obj() const noexcept { return *this; }
    ExcelObj&       obj()       noexcept { return *this; }
  };


  /// <summary>
  /// An implementation of Range which uses an ExcelRef, i.e. an Xll sheet 
  /// reference as it's underlying type
  /// </summary>
  class XllRange : public Range
  {
  public:
    explicit XllRange(const ExcelRef& ref) noexcept : _ref(ref) {}
    explicit XllRange(ExcelRef&& ref)      noexcept : _ref(ref) {}
    explicit XllRange(const ExcelObj& ref) noexcept : _ref(ExcelRef(ref)) {}

    Range* range(
      int fromRow, int fromCol,
      int toRow = TO_END, int toCol = TO_END) const final override
    {
      return new XllRange(_ref.range(fromRow, fromCol, toRow, toCol));
    }

    Range* trim() const final override
    {
      auto val = _ref.value();
      if (!val.isType(ExcelType::Multi))
        return new XllRange(*this);
      ExcelArray array(val);
      return range(0, 0, array.nRows(), array.nCols());
    }

    std::tuple<row_t, col_t> shape() const final override
    {
      return { _ref.nRows(), _ref.nCols() };
    }

    std::tuple<row_t, col_t, row_t, col_t> bounds() const final override
    {
      return _ref.bounds();
    }

    /// <summary>
    /// Returns the address of the range in the form
    /// 'SheetNm!A1:Z5'
    /// </summary>
    std::wstring address(bool local = false) const final override
    {
      return _ref.address(local);
    }

    /// <summary>
    /// Converts the referenced range to an ExcelObj. 
    /// References to single cells return an ExcelObj of the
    /// appropriate type. Multicell refernces return an array.
    /// </summary>
    ExcelObj value() const final override
    {
      return _ref.value();
    }

    ExcelObj value(row_t i, col_t j) const final override
    {
      return _ref.range(i, j, i + 1, j + 1).value();
    }

    void set(const ExcelObj& value) final override
    {
      _ref.set(value);
    }

    /// <summary>
    /// Sets the formula if the range is a cell or an array formula for a 
    /// larger range. Formulae must use RC-style references; this is not
    /// the case for ExcelRange, so there is no setFormula on the base Range
    /// class 
    /// <see cref="xloil::Range"\> class.
    /// </summary>
    void setFormula(const std::wstring_view& formula);

    std::wstring formula() final override
    {
      // xlfGetFormula always returns RC references, but GetCell uses the
      // workspace settings to return RC or A1 style.
      return callExcel(msxll::xlfGetCell, 6, _ref).get<std::wstring>();
    }

    /// <summary>
    /// Clears / empties all cells referred to by this ExcelRange.
    /// </summary>
    void clear() final
    {
      _ref.clear();
    }

    const ExcelRef& asRef() const { return _ref; }

  private:
    ExcelRef _ref;
  };
}