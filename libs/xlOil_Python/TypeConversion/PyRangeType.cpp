#include "PyCore.h"
#include "PyHelpers.h"
#include "BasicTypes.h"
#include <xlOil/ExcelRange.h>

using std::shared_ptr;
using std::vector;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      class PyFromRange : public PyFromExcelImpl
      {
      public:
        using PyFromExcelImpl::operator();
        static constexpr char* const ourName = "Range";

        PyObject* operator()(RefVal obj) const 
        {
          return pybind11::cast(newXllRange(obj)).release().ptr();
        }
        constexpr wchar_t* failMessage() const { return L"Expected range"; }
      };
      static int theBinder = addBinder([](pybind11::module& mod)
      {
        bindPyConverter<PyFromExcelConverter<PyFromRange>>(mod, "Range").def(py::init<>());
      });
    }
  }
}