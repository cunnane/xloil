#pragma once
#include <xlOil/ExcelRange.h>
#include <comip.h>

namespace Excel {
  struct __declspec(uuid("00020846-0000-0000-c000-000000000046")) Range;
}

namespace xloil
{
  namespace COM
  {
    /// <summary>
    /// This is identical to an Excel::RangePtr, but I don't know how to fwd declare it
    /// </summary>
    using OurRangePtr = _com_ptr_t<_com_IIID<Excel::Range, &__uuidof(Excel::Range)> >;

    /// <summary>
    /// An implemention of Range based which uses Excel::Range as it's 
    /// underlying type. Used when the XLL interface is not safely available.
    /// </summary>
    class ComRange : public Range
    {
    public:
      explicit ComRange(const wchar_t* address);
      ComRange(const OurRangePtr& range);

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

      const Excel::Range& native() const { return _range; }

    private:
      OurRangePtr _range;
    };
  }
}