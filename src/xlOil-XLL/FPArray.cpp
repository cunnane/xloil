#include <xloil/FPArray.h>
#include <xlOil/Events.h>
#include <xlOil/Throw.h>
#include <mutex>

using std::vector;
using std::shared_ptr;

namespace xloil
{
  namespace
  {
    auto createArray(size_t nRows, size_t nCols)
    {
      if (nRows == 0 || nCols == 0)
        XLO_THROW("Cannot create an empty FPArray");

      auto n = nRows * nCols;
      
      shared_ptr<FPArray> fp(
        (FPArray*)new char[sizeof(msxll::_FP12) + n * sizeof(double)],
        [](auto* p) { delete[] p; });
      assert(fp.get());

      fp->rows = (int)nRows;
      fp->columns = (int)nCols;
      return fp;
    }

    static vector<shared_ptr<FPArray>> theAliveArrays;
    static std::mutex theAliveArraysLock;

    // No need for lock as events are single threaded
    static auto handler = Event::AfterCalculate() += []() { theAliveArrays.clear(); };

    static auto theEmptyArray = []() {
      auto fp = createArray(1, 1);
      (*fp)(0, 0) = std::numeric_limits<double>::quiet_NaN();
      return fp;
    }();
  }

  shared_ptr<FPArray> FPArray::create(size_t nRows, size_t nCols)
  {
    auto fp = createArray(nRows, nCols);
    std::lock_guard lock(theAliveArraysLock);
    theAliveArrays.push_back(fp);
    return fp;
  }

  FPArray* FPArray::empty()
  {
    return theEmptyArray.get();
  }
}