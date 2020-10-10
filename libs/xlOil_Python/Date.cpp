#include "xloil/Date.h"
#include "BasicTypes.h"
#include "PyCoreModule.h"
#include <Python.h>
#include <datetime.h>

namespace py = pybind11;

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

    class PyFromDate : public FromExcelBase<PyObject*>
    {
    public:
      using FromExcelBase::operator();

      PyObject* operator()(int x) const 
      {
        int day, month, year;
        excelSerialDateToYMD(x, year, month, day);
        return PyDate_FromDate(year, month, day);
      }
      PyObject* operator()(double x) const
      {
        return operator()(int(x));
      }
      PyObject* operator()(const PStringView<>& pstr) const
      {
        std::tm tm;
        if (stringToDateTime(pstr.view(), tm))
          return PyDate_FromDate(tm.tm_year, tm.tm_mon, tm.tm_yday);
        return nullptr;
      }
      constexpr wchar_t* failMessage() const { return L"Expected date"; }
    };

    class PyFromDateTime : public FromExcelBase<PyObject*>
    {
    public:
      using FromExcelBase::operator();

      PyObject* operator()(int x) const
      {
        return PyFromDate()(x);
      }

      PyObject* operator()(double x) const
      {
        int day, month, year, hours, mins, secs, usecs;
        excelSerialDatetoYMDHMS(x, year, month, day, hours, mins, secs, usecs);
        return PyDateTime_FromDateAndTime(year, month, day, hours, mins, secs, usecs);
      }

      PyObject* operator()(const PStringView<>& pstr) const
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
    };
    namespace
    {
      template<class T>
      void declare2(pybind11::module& mod, const char* name)
      {
        py::class_<T, IPyToExcel, shared_ptr<T>>(mod, name)
          .def(py::init<>());
      }
      static int theBinder = addBinder([](py::module& mod)
      {
        bindPyConverter<PyExcelConverter<PyFromDateTime>>(mod, "datetime").def(py::init<>());
        bindPyConverter<PyExcelConverter<PyFromDate>>(mod, "date").def(py::init<>());
        bindXlConverter<PyDateTimeToExcel>(mod, "datetime").def(py::init<>());
        bindXlConverter<PyDateToExcel>(mod, "date").def(py::init<>());
      });
    }
  }
}