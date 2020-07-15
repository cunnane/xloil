#include "PyHelpers.h"

namespace xloil {
  namespace Python
  {
    bool sliceHelper(
      const pybind11::tuple& loc,
      const size_t nRows, const size_t nCols,
      size_t& fromRow, size_t& fromCol,
      size_t& toRow, size_t& toCol)
    {
      if (loc.size() != 2)
        XLO_THROW("Expecting tuple of size 2");
      auto r = loc[0];
      auto c = loc[1];
      size_t step = 1;
      bool singleElement = false;

      if (r.is_none())
      {
        fromRow = 0;
        toRow = nRows;
      }
      else if (PySlice_Check(r.ptr()))
      {
        size_t sliceLength;
        r.cast<pybind11::slice>().compute(nRows, &fromRow, &toRow, &step, &sliceLength);
      }
      else
      {
        fromRow = r.cast<size_t>();
        toRow = fromRow + 1;
        singleElement = true;
      }

      if (c.is_none())
      {
        fromCol = 0;
        toCol = nCols;
      }
      else if (PySlice_Check(c.ptr()))
      {
        size_t sliceLength;
        c.cast<pybind11::slice>().compute(nCols, &fromCol, &toCol, &step, &sliceLength);
      }
      else
      {
        fromCol = c.cast<size_t>();
        if (singleElement)
          return true;
        toCol = fromCol + 1;
      }

      if (step != 1)
        XLO_THROW("Slice step size must be 1");

      return false;
    }
  }
}