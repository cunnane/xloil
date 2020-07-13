#include "File.h"

#include <COMInterface/ExcelTypeLib.h>
#include "PyHelpers.h"
#include "FunctionRegister.h"
#include "Main.h"

#include <xloil/Log.h>
#include <xloil/State.h>
#include <filesystem>

namespace fs = std::filesystem;

using std::vector;
using std::string;
using std::wstring;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    void scanModule(const py::object& mod, const wchar_t* workbookName)
    {
      py::gil_scoped_acquire get_gil;

      const auto xloilModule = py::module::import("xloil");
      const auto scanFunc = xloilModule.attr("scan_module").cast<py::function>();

      const auto wbName = workbookName 
        ? py::wstr(workbookName) 
        : py::none();

      const auto modName = (string)py::str(mod);
      try
      {
        XLO_INFO("Scanning module {0}", modName);
        scanFunc.call(mod, wbName);
      }
      catch (const std::exception& e)
      {
        auto pyPath = (string)py::str(PyBorrow<py::list>(PySys_GetObject("path")));
        XLO_ERROR("Error reading module {0}: {1}\nsys.path={2}", 
          modName, e.what(), pyPath);
      }
    }

    struct WorkbookOpenHandler
    {
      WorkbookOpenHandler(const wstring& starredPattern)
      {
        // Turn the starred pattern into a fmt string for easier substitution later
        _workbookPattern = starredPattern;
        _workbookPattern.replace(_workbookPattern.find(L'*'), 1, wstring(L"{0}\\{1}"));
      }

      wstring _workbookPattern;

      void operator()(const wchar_t* wbPath, const wchar_t* wbName) const
      {
        // Subtitute in to find target module name, removing extension
        auto modulePath = fmt::format(_workbookPattern,
          wbPath,
          wstring(wbName, wcsrchr(wbName, L'.')));

        if (!fs::exists(modulePath))
          return;
        try
        {
          // First add the module, if the scan fails it will still be on the
          // file change watchlist
          FunctionRegistry::addModule(theCoreContext, modulePath, wbName);
          scanModule(py::wstr(modulePath), wbName);
        }
        catch (const std::exception& e)
        {
          XLO_WARN(L"Failed to load module {0}: {1}", modulePath, utf8ToUtf16(e.what()));
        }
      }
    };

    void checkWorkbooksOnOpen(const WorkbookOpenHandler& handler)
    {
      try
      {
        auto& workbooks = State::excelApp().Workbooks;
        auto nWorkbooks = workbooks->Count;
        for (auto i = 0; i < nWorkbooks; ++i)
        {
          auto wb = workbooks->GetItem(_variant_t(i + 1));
          handler(wb->Path, wb->Name);
        }
      }
      XLO_RETHROW_COM_ERROR;
    }

    void createWorkbookOpenHandler(const wchar_t* starredPattern)
    {
      if (!wcschr(starredPattern, L'*'))
      {
        XLO_WARN("WorkbookModule should be of the form '*foo.py' where '*'"
          "will be replaced by the full workbook path with file extension removed");
        return;
      }
      WorkbookOpenHandler handler(starredPattern);

      checkWorkbooksOnOpen(handler);

      static auto wbOpenHandler = Event::WorkbookOpen().bind(handler);
    }
  }
}