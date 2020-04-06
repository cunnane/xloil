#include "InjectedModule.h"
#include "PyHelpers.h"
#include "PyExcelArray.h"
#include "BasicTypes.h"
#include <xlOil/ExcelRange.h>
#include <xloil/Log.h>

using std::shared_ptr;
using std::vector;
namespace py = pybind11;

namespace xloil {
  namespace Python {

    using BinderFunc = std::function<void(pybind11::module&)>;
    void bindFirst(py::module& mod);
    namespace
    {
      class BinderRegistry
      {
      public:
        static BinderRegistry& get() {
          static BinderRegistry instance;
          return instance;
        }

        void add(BinderFunc f)
        {
          theFunctions.push_back(f);
        }

        void bindAll(py::module& mod)
        {
          bindFirst(mod);
          for (auto m : theFunctions)
            m(mod);
        }
      private:
        BinderRegistry() {}
        vector<BinderFunc> theFunctions;
      };
    }

    PyObject* buildInjectedModule()
    {
      auto mod = py::module(theInjectedModuleName);
      BinderRegistry::get().bindAll(mod);
      return mod.release().ptr();
    }

    int addBinder(std::function<void(pybind11::module&)> binder)
    {
      BinderRegistry::get().add(binder);
      return 0;
    }

    void scanModule(py::object& mod)
    {
      py::gil_scoped_acquire get_gil;

      auto oilModule = py::module::import("xloil");
      auto scanFunc = oilModule.attr("scan_module").cast<py::function>();
            
      try
      {
        XLO_INFO("Scanning module {0}", (std::string)py::str(mod));
        scanFunc.call(mod);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Error reading module {0}: {1}", (std::string)py::str(mod) , e.what());
      }
    }

    PyTypeObject* pyExcelErrorType = nullptr;
    PyTypeObject* ExcelArrayType = nullptr;
    PyTypeObject* ExcelRangeType = nullptr;

    inline auto subrange(const ExcelRange& r,
      int fromR, int fromC, int* nRows = 0, int* nCols = 0, int* toR = 0, int* toC = 0)
    {
      if (!(toR || nRows))
        XLO_THROW("Must specify end row or number of rows");
      if (!(toC || nCols))
        XLO_THROW("Must specify end column or number of columns");
      return r.range(fromR, fromC,
        toR ? *toR : fromR + *nRows,
        toC ? *toC : fromC + *nCols);
    }

    auto toArray(const PyExcelArray& arr, std::optional<int> dtype, std::optional<int> dims)
    {
      return PySteal<>(excelArrayToNumpyArray(arr.base(), dims ? *dims : 2, dtype ? *dtype : -1));
    }
    void bindFirst(py::module& mod)
    {
      py::class_<IPyFromExcel, shared_ptr<IPyFromExcel>>(mod, "IPyFromExcel")
        .def("__call__",
          [](const IPyFromExcel& self, const py::object& arg)
      {
        if (Py_TYPE(arg.ptr()) == ExcelArrayType)
        {
          auto arr = arg.cast<PyExcelArray>();
          return self.fromArray(arr.base());
        }
        else if (PyLong_Check(arg.ptr()))
        {
          return self(ExcelObj(arg.cast<long>()));
        }
        XLO_THROW("Not implemented");
      });
      py::class_<IPyToExcel, shared_ptr<IPyToExcel>>(mod, "IPyToExcel");

      // Bind CellError type to xloil::CellError enum
      auto eType = py::enum_<CellError>(mod, "CellError");
      for (auto e : theCellErrors)
        eType.value(utf16ToUtf8(enumAsWCString(e)).c_str(), e);

      pyExcelErrorType = (PyTypeObject*)eType.get_type().ptr();

      // Bind Range class from xloil::ExcelRange
      auto rType = py::class_<ExcelRange>(mod, "Range")
        .def("range", subrange,
          py::arg("from_row"), py::arg("from_col"),
          py::arg("num_rows") = -1, py::arg("num_cols") = -1,
          py::arg("to_row") = nullptr, py::arg("to_col") = nullptr)
        .def("cell", &ExcelRange::cell, py::arg("row"), py::arg("col"))
        .def_property("value",
          [](const ExcelRange& r)
          {
            return PySteal<>(PyFromExcel<PyFromAny<>>()(r.value()));
          },
          [](ExcelRange& r, const py::object& value)
          {
            r = FromPyObj()(value.ptr());
          },
        py::return_value_policy::automatic)
        .def("set",
          [](ExcelRange& r, const py::object& value)
          {
            r = FromPyObj()(value.ptr());
          })
        .def("clear", &ExcelRange::clear)
        .def("address", &ExcelRange::address, py::arg("local") = false)
        .def_property_readonly("nrows", &ExcelRange::nRows)
        .def_property_readonly("ncols", &ExcelRange::nCols);

      auto aType = py::class_<PyExcelArray>(mod, "ExcelArray")
        .def("sub_array", &PyExcelArray::subArray, py::arg("from_row"), py::arg("from_col"),
          py::arg("to_row") = 0, py::arg("to_col") = 0)
        .def("to_numpy", &toArray,
          py::arg("dtype")=py::none(), py::arg("dims")=2)
        .def("__call__", 
          [](const PyExcelArray& self, int i, int j)
          {
            return self(i, j);
          })
        .def("__getitem__", &PyExcelArray::getItem)
        .def_property_readonly("nrows", &PyExcelArray::nRows)
        .def_property_readonly("ncols", &PyExcelArray::nCols)
        .def_property_readonly("dims", &PyExcelArray::dims);

      ExcelArrayType = (PyTypeObject*)rType.get_type().ptr();

      mod.def("to_array", &toArray,
        py::arg("array"), py::arg("dtype")=py::none(), py::arg("dims")=2);


    }
} }