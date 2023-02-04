#include "PyAddin.h"
#include "EventLoop.h"
#include "PyCore.h"
#include "PyFunctionRegister.h"
#include "PyHelpers.h"
#include <xlOil/Interface.h>
#include <xlOil/DynamicRegister.h>
#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <toml++/toml.h>
#include <pybind11/stl.h>
#include <datetime.h> // From CPython

using std::vector;
using std::wstring;
using std::string;
using std::shared_ptr;
using std::make_shared;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      std::map<wstring, std::shared_ptr<PyAddin>> theAddins;
    }

    PyAddin::PyAddin(AddinContext& ctx, bool newThread, const wchar_t* wbPattern)
      : context(ctx)
      , thread(newThread ? make_shared<EventLoop>() : theCoreAddin()->thread)
    {
      if (wbPattern)
      {
        _workbookPattern = wbPattern;
        const auto star = _workbookPattern.find(L'*');
        if (star == wstring::npos)
        {
          XLO_WARN("WorkbookModule should be of the form '*foo.py' where '*'"
            "will be replaced by the full workbook path with file extension removed");
        }
        else // Replace the star so we can use formatStr later
          _workbookPattern.replace(star, 1, wstring(L"%s"));
      }
    }

    const std::wstring& PyAddin::pathName() const
    {
      return context.pathName();
    }

    std::wstring PyAddin::getLocalModulePath(const wchar_t* workbookPath) const
    {
      // Substitute in to find target module name, removing extension
      auto fileExtn = wcsrchr(workbookPath, L'.');
      return formatStr(_workbookPattern.c_str(),
        fileExtn ? wstring(workbookPath, fileExtn).c_str() : workbookPath);
    }

    void PyAddin::importModule(const pybind11::object& module)
    {
      return thread->callback("xloil.importer", "_import_and_scan",
        module, self());
    }

    void PyAddin::importFile(const wchar_t* filePath, const wchar_t* linkedWorkbook)
    {
      return thread->callback("xloil.importer", "_import_file_and_scan",
        filePath, self(), linkedWorkbook);
    }

    pybind11::object PyAddin::self() const
    {
      return py::cast(shared_from_this());
    }

    std::shared_ptr<FuncSource> PyAddin::findSource(const wchar_t* sourcePath) const
    {
      auto found = context.sources().find(sourcePath);
      if (found != context.sources().end())
        return found->second;
      return std::shared_ptr<FuncSource>();
    }

    PyAddin& findAddin(const wchar_t* xllPath)
    {
      const auto found = xllPath ? theAddins.find(xllPath) : theAddins.end();
      if (found == theAddins.end())
        XLO_THROW(L"Could not find python addin for {}", xllPath);
      return *found->second;
    }
    
    std::map<wstring, std::shared_ptr<PyAddin>>& getAddins()
    {
      return theAddins;
    }

    std::shared_ptr<EventLoop> getEventLoop()
    {
      const auto id = std::this_thread::get_id();
      for (auto& [key, addin] : theAddins)
        if (addin->thread->thread().get_id() == id)
          return addin->thread;
      XLO_THROW("Internal: could not find addin associated with current thread");
    }

    auto findAllAddinFuncs(PyAddin& addin)
    {
      vector<shared_ptr<PyFuncInfo>> funcInfo;
      for (auto&[name, source] : addin.context.sources())
      {
        auto pySource = std::dynamic_pointer_cast<RegisteredModule>(source);
        if (!pySource)
          continue; // Unexpected - give warning?

        for (auto& funcSpec : pySource->functions())
        {
          auto pySpec = std::static_pointer_cast<const DynamicSpec>(funcSpec);
          auto pyFuncInfo = std::static_pointer_cast<const PyFuncInfo>(pySpec->context());
          funcInfo.push_back(std::const_pointer_cast<PyFuncInfo>(pyFuncInfo));
        }
      }
      return funcInfo;
    }

    // TODO: replace with pybind11::bind_map
    template<class TMap>
    class PyWrapMap
    {
    private:
      const TMap& _map;

    public:
      PyWrapMap(const TMap& mapRef) : _map(mapRef) {}

      auto keys() const
      {
        return py::make_key_iterator(_map.begin(), _map.end());
      }
      auto values() const
      {
        return py::make_value_iterator(_map.begin(), _map.end());
      }
      auto items() const
      {
        return py::make_iterator(_map.begin(), _map.end());
      }

      auto getItem(const typename TMap::key_type& key) const
      {
        auto found = _map.find(key);
        if (found == _map.end())
          throw py::key_error();
        return found->second;
      }
      size_t len() const { return _map.size(); }

      bool contains(const typename TMap::key_type& key) const
      {
        return _map.find(key) != _map.end();
      }

      using this_t = PyWrapMap<TMap>;
      static void bind(py::module& mod, const char* name)
      {
        py::class_<this_t>(mod, name, 
           R"(
             A dictionary of all addins using the xlOil_Python plugin keyed
             by the addin pathname.
           )")
          .def("__getitem__", &this_t::getItem)
          .def("__len__", &this_t::len)
          .def("__iter__", &this_t::keys)
          .def("__contains__", &this_t::contains)
          .def("items", &this_t::items)
          .def("values", &this_t::values)
          .def("keys", &this_t::keys);
      }
    };

    py::object tomlNodeToPyObject(const toml::node& node)
    {
      using toml::node_type;
      
      switch (node.type())
      {
      case node_type::table:          return py::cast(py::ReferenceHolder(node.as_table()));
      case node_type::string:         return py::cast(**node.as_string());
      case node_type::integer:        return py::cast(**node.as_integer());
      case node_type::floating_point: return py::cast(**node.as_floating_point());
      case node_type::boolean:        return py::cast(**node.as_boolean());
      case node_type::none:           return py::none();
      case node_type::date:
      {
        const auto& date = **node.as_date();
        return PySteal<>(PyDate_FromDate(date.year, date.month, date.day));
      }
      case node_type::time:
      {
        const auto& time = **node.as_time();
        return PySteal<>(PyTime_FromTime(time.hour, time.minute, time.second, time.nanosecond * 100));
      }
      case node_type::date_time:
      {
        const auto& datetime = **node.as_date_time();
        return PySteal<>(PyDateTime_FromDateAndTime(
          datetime.date.year, datetime.date.month, datetime.date.day,
          datetime.time.hour, datetime.time.minute, datetime.time.second,
          datetime.time.nanosecond * 100));
      }
      
      case node_type::array:
      {
        const auto& array = *node.as_array();
        auto list = py::list(array.size());
        for (size_t i = 0; i < array.size(); ++i)
          list[i] = tomlNodeToPyObject(array[i]);
        return list;
      }
      default:
        // We support all types as of Sept 2022, so if we get here something 
        // was corrupted or a new type has been added..
        throw py::type_error("Unsupported toml node type");
      }
    }
    auto tomlTableGetItem(toml::table& table, const char* name)
    {
      const auto* node = table.get(name);
      if (!node)
        throw py::key_error(name);
      return tomlNodeToPyObject(*node);
    }
    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<toml::table, py::ReferenceHolder<toml::table>>(mod, "_TomlTable")
          .def("__getitem__", tomlTableGetItem);

        py::class_<PyAddin, shared_ptr<PyAddin>>(mod, "Addin")
          .def_property_readonly("pathname", &PyAddin::pathName)
          .def_property_readonly("event_loop",
            [](PyAddin& addin) { return addin.thread->loop(); },
            R"(
              The asyncio event loop used for background tasks by this addin
            )")
          .def_property_readonly("settings_file",
            [](PyAddin& addin) { return string(* addin.context.settings()->source().path); },
            R"(
              The full pathname of the settings ini file used by this addin
            )")
          .def_property_readonly("settings",
            [](PyAddin& addin) { return py::cast(py::ReferenceHolder(addin.context.settings())); },
            R"(
              Gives access to the settings in the addin's ini file as nested dictionaries.
              These are the settings on load and do not allow for modifications made in the 
              ribbon toolbar.
            )")
          .def("functions", findAllAddinFuncs,
            R"(
              Returns a list of all functions declared by this addin.
            )")
          .def("source_files",
            [](PyAddin& addin)
            {
              vector<wstring> sources;
              for (auto& item : addin.context.sources())
                sources.push_back(item.first);
              return sources;
            });

        PyWrapMap<decltype(theAddins)>::bind(mod, "_AddinsDict");

        mod.def("core_addin", theCoreAddin);
        mod.add_object("xloil_addins", 
          py::cast(new PyWrapMap(theAddins), py::return_value_policy::take_ownership));
      });
    }
  }
}