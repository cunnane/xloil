#include "NumpyHelpers.h"
#include "PyCore.h"
#include "BasicTypes.h"

using row_t = xloil::ExcelArray::row_t;
using col_t = xloil::ExcelArray::col_t;
namespace py = pybind11;
using std::vector;
using std::unique_ptr;

namespace xloil
{
  namespace Python
  {
    namespace TableHelpers
    {
      struct ApplyConverter
      {
        virtual ~ApplyConverter() {}
        virtual void operator()(ExcelArrayBuilder& builder,
          xloil::detail::ArrayBuilderIterator& start,
          xloil::detail::ArrayBuilderIterator& end) = 0;
      };

      template<int NPDtype>
      struct ConverterHolder : public ApplyConverter
      {
        FromArrayImpl<NPDtype> _impl;
        PyArrayObject* _array;

        ConverterHolder(PyArrayObject* array, bool)
          : _impl(array)
          , _array(array)
        {}

        virtual ~ConverterHolder() {}

        size_t stringLength() const { return _impl.stringLength; }

        virtual void operator()(ExcelArrayBuilder& builder,
          xloil::detail::ArrayBuilderIterator& start,
          xloil::detail::ArrayBuilderIterator& end)
        {
          char* arrayPtr = PyArray_BYTES(_array);
          const auto step = PyArray_STRIDE(_array, 0);
          for (; start != end; arrayPtr += step, ++start)
          {
            start->take(_impl.toExcelObj(builder, arrayPtr));
          }
        }
      };

      template<>
      struct ConverterHolder<NPY_OBJECT> : public ApplyConverter
      {
        SequentialArrayBuilder _builder;

        ConverterHolder(PyArrayObject* array, bool objectToString)
          : _builder((row_t)PyArray_DIMS(array)[0], 1)
        {
          auto arrayPtr = PyArray_BYTES(array);
          const auto N = _builder.nRows();
          const auto step = PyArray_STRIDE(array, 0);
          auto charAllocator = _builder.charAllocator();
          if (objectToString)
          {
            for (auto i = 0u; i < N; ++i, arrayPtr += step)
              _builder.emplace(
                FromPyObj<detail::ReturnToString, true>()(
                  *(PyObject**)arrayPtr, charAllocator));
          }
          else
          {
            for (auto i = 0u; i < N; ++i, arrayPtr += step)
              _builder.emplace(
                FromPyObj<detail::ReturnToCache, true>()(
                  *(PyObject**)arrayPtr, charAllocator));
          }
        }

        virtual ~ConverterHolder() {}

        size_t stringLength() const { return _builder.stringLength(); }

        virtual void operator()(ExcelArrayBuilder& /*builder*/,
          xloil::detail::ArrayBuilderIterator& start,
          xloil::detail::ArrayBuilderIterator& end)
        {
          _builder.copyToBuilder(start, end);
        }
      };

      /// <summary>
      /// Helper class used with `switchDataType`
      /// </summary>
      template<int NPDtype>
      struct CreateConverter
      {
        ApplyConverter* operator()(PyArrayObject* array, size_t& stringLength, bool objectToString)
        {
          auto converter = new ConverterHolder<NPDtype>(array, objectToString);
          stringLength += converter->stringLength();
          return converter;
        }
      };

      size_t arrayShape(const py::handle& p)
      {
        if (p.is_none())
          return 0;

        if (!PyArray_Check(p.ptr()))
          XLO_THROW("Expected an array");

        auto pyArr = (PyArrayObject*)p.ptr();
        const auto nDims = PyArray_NDIM(pyArr);
        const auto dims = PyArray_DIMS(pyArr);

        if (nDims != 1)
          XLO_THROW("Expected 1 dim array");

        return dims[0];
      }

      /// <summary>
      /// This class holds an array of virtual FromArrayImpl holders.  Each column in a 
      /// dataframe can have a different data type and so require a different converter.
      /// The indices can also have their own data types. The class uses `collect` to 
      /// examine 1-d numpy arrays and creates an appropriate converters. Then `write` is
      /// called when and ExcelArrayBuilder object is ready to receive the converted data
      /// </summary>
      struct Converters
      {
        vector<unique_ptr<ApplyConverter>> _converters;
        size_t stringLength = 0;
        bool _hasObjectDtype;
        bool _objectToString;

        Converters(size_t n, bool objectToString)
          : _objectToString(objectToString)
          , _hasObjectDtype(false)
        {
          _converters.reserve(n);
        }

        auto collect(const py::handle& p, size_t expectedLength)
        {
          auto shape = arrayShape(p);

          if (shape != expectedLength)
            XLO_THROW("Expected a 1-dim array of size {}", expectedLength);

          auto pyArr = (PyArrayObject*)p.ptr();
          const auto dtype = PyArray_TYPE(pyArr);
          if (dtype == NPY_OBJECT)
            _hasObjectDtype = true;

          _converters.emplace_back(unique_ptr<ApplyConverter>(
            switchDataType<CreateConverter>(dtype, pyArr, std::ref(stringLength), _objectToString)));
        }

        auto write(size_t iArray, ExcelArrayBuilder& builder, int startX, int startY, bool byRow)
        {
          auto start = byRow
            ? builder.row_begin(startX) + startY
            : builder.col_begin(startX) + startY;

          auto end = byRow
            ? builder.row_end(startX)
            : builder.col_end(startX);

          (*_converters[iArray])(builder, start, end);
        }

        /// <summary>
        /// Used to determine if we can release the GIL for the duration of the conversion
        /// </summary>
        auto hasObjectDtype() const { return _hasObjectDtype; }
      };
    }

    ExcelObj numpyTableHelper(
      uint32_t nOuter,
      uint32_t nInner,
      const py::object& columns,
      const py::object& rows,
      const py::object& headings,
      const py::object& index,
      const py::object& indexName,
      bool useObjectCache)
    {
      // This method can handle writing data vertically or horizontally.  When used to 
      // write a pandas DataFrame, the data is vertical/by-column.
      const auto byRow = columns.is_none();

      const auto hasHeadings = !headings.is_none();
      const auto hasIndex = !index.is_none();

      auto tableData = const_cast<PyObject*>(byRow ? rows.ptr() : columns.ptr());

      // The row or column headings can be multi-level indices. We determine the number
      // of levels from iterators later.
      auto nHeadingLevels = 0u;
      auto nIndexLevels = 0u;


      // Converters may end up larger if we have multi-level indices
      TableHelpers::Converters converters(
        nOuter + (hasHeadings ? 1 : 0) + (hasIndex ? 1 : 0), !useObjectCache);

      // Examine data frame index
      if (hasIndex)
      {
        for (auto iter = py::iter(index); iter != py::iterator::sentinel(); ++iter)
        {
          converters.collect(*iter, nInner);
          ++nIndexLevels;
        }
      }

      
      // First loop to establish array size and length of strings
      for (auto iter = py::iter(tableData); iter != py::iterator::sentinel(); ++iter)
      {
        converters.collect(*iter, nInner);
      }

      if (hasHeadings)
      {
        for (auto iter = py::iter(headings); iter != py::iterator::sentinel(); ++iter)
        {
          converters.collect(*iter, nOuter);
          ++nHeadingLevels;
        }
      }

      vector<ExcelObj> indexNames(nIndexLevels * nHeadingLevels, CellError::NA);
      auto indexNameStringLength = 0;
      if (nIndexLevels > 0 && !indexName.is_none())
      {
        auto i = 0u;
        for (auto iter = py::iter(indexName); i < nIndexLevels * nHeadingLevels && iter != py::iterator::sentinel(); ++i, ++iter)
        {
          indexNames[i] = FromPyObj()(iter->ptr());
          indexNameStringLength += indexNames[i].stringLength();
        }
      }

      // If possible, release the GIL before beginning the conversion
      NumpyBeginThreadsDescr releaseGil(
        converters.hasObjectDtype() ? NPY_OBJECT : NPY_FLOAT);

      auto nRows = nOuter + nIndexLevels;
      auto nCols = nInner + nHeadingLevels;
      if (!byRow)
        std::swap(nRows, nCols);


      ExcelArrayBuilder builder(
        nRows,
        nCols,
        converters.stringLength + indexNameStringLength);

      // Write the index names in the top left
      if (!byRow)
      {
        auto k = 0u;
        for (auto i = 0u; i < nHeadingLevels; ++i)
          for (auto j = 0u; j < nIndexLevels; ++j)
            builder(i, j) = indexNames[k++];
      }
      else
      {
        auto k = 0u;
        for (auto j = 0u; j < nIndexLevels; ++j)
          for (auto i = 0u; i < nHeadingLevels; ++i)
            builder(i, j) = indexNames[k++];
      }

      auto iConv = 0;

      for (auto i = 0u; i < nOuter + nIndexLevels; ++i, ++iConv)
        converters.write(iConv, builder, i, nHeadingLevels, byRow);

      for (auto i = 0u; i < nHeadingLevels; ++i, ++iConv)
        converters.write(iConv, builder, i, nIndexLevels, !byRow);

      return builder.toExcelObj();
    }
    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        mod.def("_table_converter",
          &numpyTableHelper,
          R"(
          For internal use. Converts a table like object (such as a pandas DataFrame) to 
          RawExcelValue suitable for returning to xlOil.
            
          n, m:
            the number of data fields and the length of the fields
          columns / rows: 
            a iterable of numpy array containing data, specified as columns 
            or rows (not both)
          headings:
            optional array of data field headings
          index:
            optional data field labels - one per data point
          index_name:
            optional headings for the index, should be a 1 dim iteratable of size
            num_index_levels * num_column_levels
          cache_objects:
            if True, place unconvertible objects in the cache and return a ref string
            if False, call str(x) on unconvertible objects
          )",
          py::arg("n"),
          py::arg("m"),
          py::arg("columns") = py::none(),
          py::arg("rows") = py::none(),
          py::arg("headings") = py::none(),
          py::arg("index") = py::none(),
          py::arg("index_name") = py::none(),
          py::arg("cache_objects") = false);
      });
    }
  }
}
