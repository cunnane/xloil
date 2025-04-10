#include "PyCore.h"
#include "PyHelpers.h"
#include <xlOil/ExcelRef.h>

using std::shared_ptr;
using std::vector;
using std::wstring;

namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    AddressStyle parseAddressStyle(const std::string_view& style)
    {
      if (style == "" || style == "a1")
        return AddressStyle::A1;
      else if (style == "$a1")
        return AddressStyle::A1 | AddressStyle::COL_FIXED;
      else if (style == "a$1")
        return AddressStyle::A1 | AddressStyle::ROW_FIXED;
      else if (style == "$a$1")
        return AddressStyle::A1 | AddressStyle::ROW_FIXED | AddressStyle::COL_FIXED;
      else if (style == "rc")
        return AddressStyle::RC;
      else if (style == "$rc")
        return AddressStyle::RC | AddressStyle::ROW_FIXED;
      else if (style == "r$c")
        return AddressStyle::RC | AddressStyle::COL_FIXED;
      else if (style == "$r$c")
        return AddressStyle::RC | AddressStyle::ROW_FIXED | AddressStyle::COL_FIXED;
      
      XLO_THROW("Unknown address style '{}'", style);
    }

    namespace
    {
      class ParsedAddress
      {
      public:
        friend class ParsedAddressIter;

        ParsedAddress(
          const msxll::XLREF12& ref,
          const std::wstring& sheetName)
          : ParsedAddress(ref)
        {
          _sheetName = sheetName;
        }

        ParsedAddress(const msxll::XLREF12& ref)
          : _ref(ref)
        {
          _isRange = _ref.rwFirst != _ref.rwLast || _ref.colFirst != _ref.colLast;
        }

        ParsedAddress(const py::object& address, const py::object& sheetName)
        {
          if (PyUnicode_Check(address.ptr()))
            addressToXlRef(to_wstring(address.ptr()), _ref, &_sheetName);

          else if (PyIterable_Check(address.ptr()))
          {
            auto iter = py::iter(address);
            _ref.rwFirst = py::cast<int>(*iter++);
            _ref.colFirst = py::cast<int>(*iter++);
            if (iter != py::iterator::sentinel())
            {
              _ref.rwLast = py::cast<int>(*iter++);
              _ref.colLast = py::cast<int>(*iter++);
            }
            else
            {
              _ref.rwLast = _ref.rwFirst;
              _ref.colLast = _ref.colFirst;
            }
          }
          else
            throw py::value_error("address");

          if (!sheetName.is_none())
            _sheetName = to_wstring(sheetName);

          _isRange = _ref.rwFirst != _ref.rwLast || _ref.colFirst != _ref.colLast;
        }

        auto fromRow() const { return _ref.rwFirst; }
        auto fromCol() const { return _ref.colFirst; }
        auto toRow()   const { return _ref.rwLast; }
        auto toCol()   const { return _ref.colLast; }
        auto sheet()   const { return _sheetName; }

        auto tuple() const
        {
          if (_isRange)
            return py::make_tuple(
              _ref.rwFirst, _ref.colFirst,
              _ref.rwLast, _ref.colLast);
          else
            return py::make_tuple(_ref.rwFirst, _ref.colFirst);
        }

        auto string(AddressStyle style) const
        {
          wchar_t buf[XL_FULL_ADDRESS_RC_MAX_LEN];
          auto nWritten = xlrefToAddress(_ref, buf, sizeof(buf),
            _sheetName, style);
          return std::wstring(buf, nWritten);
        }

        auto address(std::string& style, const bool local) const
        {
          toLower(style);
          auto s = parseAddressStyle(style);
          if (local)
            s |= AddressStyle::LOCAL;
          return string(s);
        }

        auto iter() const;

      private:
        msxll::XLREF12 _ref;
        bool _isRange;
        std::wstring _sheetName;
      };

      class ParsedAddressIter
      {
      private:
        msxll::XLREF12 _parent;
        ParsedAddress _current;

      public:
        ParsedAddressIter(
          const ParsedAddress& r,
          int i, int j)
          : _parent(r._ref)
          , _current(msxll::XLREF12{ i, i, j, j }, r.sheet())
        {
        }

        // Ctor for end iterator
        ParsedAddressIter(const ParsedAddress& r)
          : _parent(r._ref)
          , _current(msxll::XLREF12{
              r.toRow() + 1, r.toRow() + 1,
              r.fromCol(), r.fromCol()
            })
        {
        }

        ParsedAddressIter& operator++()
        {
          ++_current._ref.colFirst;
          ++_current._ref.colLast;
          if (_current._ref.colFirst == _parent.colLast)
          {
            _current._ref.colFirst = _parent.colFirst;
            _current._ref.colLast = _parent.colFirst;
            ++_current._ref.rwFirst;
            ++_current._ref.rwLast;
          }

          return (*this);
        }

        const ParsedAddress& operator*() const
        {
          return _current;
        }

        bool operator==(const ParsedAddressIter& that)
        {
          return memcmp(&_current._ref, &that._current._ref,
            sizeof(msxll::XLREF12)) == 0;
        }
      };

      auto ParsedAddress::iter() const
      {
        auto begin = ParsedAddressIter(*this, fromRow(), fromCol());
        auto end = ParsedAddressIter(*this);
        return py::make_iterator(std::move(begin), std::move(end));
      }

      static int theBinder = addBinder([](pybind11::module& mod)
        {
          py::class_<ParsedAddress>(mod, "Address",
            R"(
            Converts cell addresses between different formats. This class
            only performs manipulation on the address strings, so it is fast
            and does not need to run on the main thread. However, it does not
            check the validity of the addresses. 

            Parameters
            ----------

            address: string|tuple
              either an address string in A1 format e.g. "D2:R2" or RC format, 
              e.g. "R5C3:R6C6" or a tuple of ints `(row, col)` or 
              `(from_row, from_col, to_row, to_col)`.  A string address may be
              absolute (contain dollars) and contain a workbook/sheet name, separated
              by a pling (!).  A tuple follows python conventions so is zero-based.

            sheet: [string]
              optional sheet name in case it was not passed as part of the address

            )")
            .def(py::init<const py::object&, const py::object& >(),
              py::arg("address"),
              py::arg("sheet") = py::none())
            .def("__str__", [](ParsedAddress& x) { return x.string(AddressStyle::A1); })
            .def("__iter__",
              &ParsedAddress::iter)
            .def("__call__",
              &ParsedAddress::address,
              py::arg("style") = "a1",
              py::arg("local") = false,
              R"(
              Writes the address to a string in the specified style.

              Parameters
              ----------
              style: str
                The address format: "a1" or "rc". To produce an absolute / fixed addresses
                use "$a$1", "$r$c", "$a1", "a$1", etc. depending on whether you want
                both row and column to be fixed.
              local: bool
                If True, omits sheet and workbook infomation.

              )")
            .def_property_readonly("a1",
              [](ParsedAddress& x) { return x.string(AddressStyle::A1); },
              "The address in A1 format")
            .def_property_readonly("a1_fixed",
              [](ParsedAddress& x) { return x.string(AddressStyle::A1 | AddressStyle::ROW_FIXED | AddressStyle::COL_FIXED); },
              "The absolute address in A1 format (i.e. with $s)")
            .def_property_readonly("rc",
              [](ParsedAddress& x) { return x.string(AddressStyle::RC); },
              "The address in RC format")
            .def_property_readonly("rc_fixed",
              [](ParsedAddress& x) { return x.string(AddressStyle::RC | AddressStyle::ROW_FIXED | AddressStyle::COL_FIXED); },
              "The absolute address in RC format (i.e. with $s)")
            .def_property_readonly("tuple", &ParsedAddress::tuple)
            .def_property_readonly("from_row", &ParsedAddress::fromRow)
            .def_property_readonly("from_col", &ParsedAddress::fromCol)
            .def_property_readonly("to_row", &ParsedAddress::toRow)
            .def_property_readonly("to_col", &ParsedAddress::toCol)
            .def_property_readonly("sheet", &ParsedAddress::sheet);

        });
    }
  }
}
