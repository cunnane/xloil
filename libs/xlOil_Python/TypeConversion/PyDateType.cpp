#include "BasicTypes.h"
#include "PyDateType.h"
#include "PyCore.h"
#include "PyHelpers.h"
#include <xloil/Date.h>
#include <Python.h>
#include <datetime.h>
#include <pybind11/stl_bind.h>

namespace py = pybind11;
using std::vector;
using std::wstring;



namespace xloil
{
  namespace Python
  {
    void importDatetime()
    {
      PyDateTime_IMPORT;
    }

    bool isPyDate(PyObject* p)
    {
      return (PyDate_CheckExact(p) || PyDateTime_CheckExact(p));
    }

    ExcelObj pyDateTimeToSerial(PyObject* p)
    {
      auto serial = excelSerialDateFromYMDHMS(
        PyDateTime_GET_YEAR(p), PyDateTime_GET_MONTH(p), PyDateTime_GET_DAY(p),
        PyDateTime_DATE_GET_HOUR(p), PyDateTime_DATE_GET_MINUTE(p), PyDateTime_DATE_GET_SECOND(p),
        PyDateTime_DATE_GET_MICROSECOND(p)
      );
      return ExcelObj(serial);
    }

    ExcelObj pyDateToSerial(PyObject* p)
    {
      auto serial = excelSerialDateFromYMD(
        PyDateTime_GET_YEAR(p), PyDateTime_GET_MONTH(p), PyDateTime_GET_DAY(p));
      return ExcelObj(serial);
    }

    ExcelObj pyDateToExcel(PyObject* p)
    {
      if (PyDateTime_CheckExact(p))
        return ExcelObj(pyDateTimeToSerial(p));
      else if (PyDate_CheckExact(p))
        return ExcelObj(pyDateToSerial(p));
      else
      {
        // Nil return used to indicate no conversion possible
        return ExcelObj();
      }
    }

    class PyFromDate : public detail::PyFromExcelImpl
    {
    public:
      using detail::PyFromExcelImpl::operator();
      static constexpr char* const ourName = "date";

      PyObject* operator()(int x) const
      {
        int day, month, year;
        if (!excelSerialDateToYMD(x, year, month, day))
          throw py::value_error("Number not a valid Excel serial date");
        return PyDate_FromDate(year, month, day);
      }
      PyObject* operator()(double x) const
      {
        return operator()(int(x));
      }
      PyObject* operator()(const PStringRef& pstr) const
      {
        std::tm tm;
        if (stringToDateTime(pstr.view(), tm))
          return PyDate_FromDate(tm.tm_year, tm.tm_mon, tm.tm_yday);
        return nullptr;
      }
      constexpr wchar_t* failMessage() const { return L"Expected date"; }
    };

    class PyFromDateTime : public detail::PyFromExcelImpl
    {
    public:
      using detail::PyFromExcelImpl::operator();
      static constexpr char* const ourName = "datetime";

      PyObject* operator()(int x) const
      {
        return PyFromDate()(x);
      }

      PyObject* operator()(double x) const
      {
        int day, month, year, hours, mins, secs, usecs;
        if (!excelSerialDatetoYMDHMS(x, year, month, day, hours, mins, secs, usecs))
          throw py::value_error("Number not a valid Excel serial date");
        return PyDateTime_FromDateAndTime(year, month, day, hours, mins, secs, usecs);
      }

      PyObject* operator()(const PStringRef& pstr) const
      {
        std::tm tm;
        if (stringToDateTime(pstr.view(), tm))
          return PyDateTime_FromDateAndTime(
            tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday,
            tm.tm_hour, tm.tm_min, tm.tm_sec, 0);
        return nullptr;
      }

      constexpr wchar_t* failMessage() const { return L"Expected date"; }
    };

    class PyDateToExcel : public IPyToExcel
    {
    public:
      ExcelObj operator()(const PyObject& obj) const override
      {
        return PyDate_CheckExact(&obj)
          ? ExcelObj(pyDateToSerial((PyObject*)&obj))
          : ExcelObj();
      }
      const char* name() const override
      {
        return "date";
      }
    };
    class PyDateTimeToExcel : public IPyToExcel
    {
    public:
      ExcelObj operator()(const PyObject& obj) const override
      {
        return PyDateTime_CheckExact(&obj)
          ? ExcelObj(pyDateTimeToSerial((PyObject*)&obj))
          : ExcelObj();
      }
      const char* name() const override
      {
        return "datetime";
      }
    };
  }
}

// Make vector<wstring> opaque so we can bind a reference to a vector using
// py::bind_vector. Note the opacity only affects this compliation unit.
PYBIND11_MAKE_OPAQUE(std::vector<std::wstring>);

namespace xloil
{
  namespace Python
  {
    namespace
    {
      py::object fromExcelDate(const py::object& obj)
      {
        auto p = obj.ptr();
        if (p == Py_None)
          return py::none();
        else if (PyLong_Check(p))
          return PySteal(PyFromDate()(PyLong_AsLong(p)));
        else if (PyFloat_Check(p))
          return PySteal(PyFromDateTime()(PyFloat_AS_DOUBLE(p)));
        else if (isNumpyArray(p))
          return PySteal(toNumpyDatetimeFromExcelDateArray(p));
        else if (PyUnicode_Check(p))
          return PySteal(PyFromDateTime()(FromPyString()(p).cast<PStringRef>()));
        else if (PyDateTime_Check(p))
          return obj;
        else
          throw std::invalid_argument("No conversion to date");
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        bindPyConverter<PyFromExcelConverter<PyFromDateTime>>(mod, "datetime").def(py::init<>());
        bindPyConverter<PyFromExcelConverter<PyFromDate>>(mod, "date").def(py::init<>());
        bindXlConverter<PyDateTimeToExcel>(mod, "datetime").def(py::init<>());
        bindXlConverter<PyDateToExcel>(mod, "date").def(py::init<>());

        
        mod.def("to_datetime",
          fromExcelDate,
          R"(
            Tries to the convert the given object to a `dt.date` or `dt.datetime`:

              * Numbers are assumed to be Excel date serial numbers. 
              * Strings are parsed using the current date conversion settings.
              * A numpy array of floats is treated as Excel date serial numbers and converted
                to n array of datetime64[ns].
              * `dt.datetime` is provided is simply returned.

            Raises `ValueError` if conversion is not possible.
          )");

        mod.def("from_excel_date",
          fromExcelDate,
          R"(
            Identical to `xloil.to_datetime`.
          )");

        py::bind_vector<vector<wstring>, py::ReferenceHolder<vector<wstring>>>(mod, "_DateFormatList",
          R"(
            Registers date time formats to try when parsing strings to dates.
            See `std::get_time` for format syntax.
          )");

        mod.add_object("date_formats", 
          py::cast(py::ReferenceHolder(&theDateTimeFormats())));
      });
    }
  }
}