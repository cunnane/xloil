#include "PyAddin.h"
#include "Main.h"
#include "EventLoop.h"
#include "PyCore.h"
#include <xlOil/Interface.h>
#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <tomlplusplus/toml.hpp>
#include <pybind11/stl.h>

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
        py::class_<this_t>(mod, name)
          .def("__getitem__", &this_t::getItem)
          .def("__len__", &this_t::len)
          .def("__iter__", &this_t::items)
          .def("__contains__", &this_t::contains)
          .def("items", &this_t::items)
          .def("values", &this_t::values)
          .def("keys", &this_t::keys);
      }
    };

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyAddin, shared_ptr<PyAddin>>(mod, "Addin")
          .def_property_readonly("pathname", &PyAddin::pathName)
          .def_property_readonly("event_loop",
            [](PyAddin& addin) { return addin.thread->loop(); })
          .def_property_readonly("settings_file",
            [](PyAddin& addin) { return *addin.context.settings()->source().path; })
          .def("source_files",
            [](PyAddin& addin)
            {
              vector<wstring> sources;
              for (auto& item : addin.context.sources())
                sources.push_back(item.first);
              return sources;
            });

        PyWrapMap<decltype(theAddins)>::bind(mod, "_AddinsDict");

        mod.add_object("xloil_addins", 
          py::cast(new PyWrapMap(theAddins), py::return_value_policy::take_ownership));
      });
    }
  }
}