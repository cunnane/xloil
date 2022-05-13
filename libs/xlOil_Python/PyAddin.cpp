#include "PyAddin.h"
#include "Main.h"
#include "EventLoop.h"
#include <xlOil/Interface.h>
#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>

using std::vector;
using std::wstring;
using std::string;
using std::shared_ptr;
using std::make_shared;

namespace xloil
{
  namespace Python
  {
    PyAddin::PyAddin(AddinContext& ctx, bool newThread, const wchar_t* wbPattern)
      : context(ctx)
      , thread(newThread ? make_shared<EventLoop>() : theCoreAddin().thread)
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
        module, pathName());
    }

    void PyAddin::importFile(const wchar_t* filePath, const wchar_t* linkedWorkbook)
    {
      return thread->callback("xloil.importer", "_import_file_and_scan",
        filePath, pathName(), linkedWorkbook);
    }
  }
}