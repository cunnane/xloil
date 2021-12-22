
#include "PyHelpers.h"
#include "PyCore.h"
#include "PyImage.h"
#include "PyEvents.h"
#include "Main.h"
#include <xlOil/ExcelTypeLib.h>
#include <xloil/Log.h>
#include <xloil/Throw.h>
#include <xlOil/ExcelApp.h>
#include <xlOilHelpers/Environment.h>
#include <fcntl.h>

using std::vector;
using std::string;
using std::wstring;
using std::make_pair;
using std::to_string;


namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    //
    // Some things I tried to add an image that don't work
    // 
    // Using the temp file handle directly rather than passing the filename to Python
    // requres converting the Windows handle to a C file descriptor int.  That can
    // be done with:
    // 
    //   _open_osfhandle((intptr_t)tempFileHandle, _O_APPEND);
    //   auto file = PyFile_FromFd(..., nullptr, "w", 0, nullptr, nullptr, nullptr, true);
    //   file.attr("close")();
    // 
    // However _open_osfhandle needs to be called in the same C-runtime as Python, which
    // requires calling _Py_open_osfhandle_noraise() but this function is only exposed in
    // Python >= 3.10 and is not part of the stable ABI.
    //
    // Adding a Forms.Image.1 control and setting its Picture property could avoid the temp
    // file write:
    // 
    //   auto imageObj = shapes->AddOLEObject(_variant_t(L"Forms.Image.1"), vtMissing, vtMissing, 
    //                           VARIANT_FALSE, vtMissing, vtMissing, vtMissing, 0, 0, 100, 100);
    //   MSForms::IImage* imagePtr;
    //   imageObj->OLEFormat->Object->QueryInterface(&imagePtr);
    //   imagePtr->Picture = PicturePtr(pictureFromPilImage(image));
    // 
    // Unfortunately AddOLEObject fails (for any choice of control). This may be a security 
    // issue or it may not be callable from a worksheet function.  In any case using AddPicture2
    // gives an object which behaves like a picture to the user (e.g. resize controls).
    // 

    auto writeCellImage(
      const py::object& saveFunction, 
      const py::object& size,
      const py::object& position,
      const py::object& coordinates,
      bool compress)
    {
      try
      {
        auto& app = excelApp();
        auto caller = Excel::RangePtr(app.Caller);

        // AddPicture2 takes -1 to retain the size of the existing file
        float width = -1, height = -1; 
        if (PyUnicode_Check(size.ptr()))
        {
          string sz = toLower((string)py::str(size));
          if (strcmp(sz.c_str(), "cell") == 0)
          { 
            width = float(caller->Width);
            height = float(caller->Height);
          }
          else if (strcmp(sz.c_str(), "img") == 0)
          {} // Matches the default
          else
            throw py::value_error("Size argument is invalid");
        }
        else if (!size.is_none())
        {
          auto sizePair = size.cast<py::tuple>();
          width = sizePair[0].cast<float>();
          height = sizePair[1].cast<float>();
        }

        float posX = 0, posY = 0;
        if (!position.is_none())
        {
          auto pos = position.cast<py::tuple>();
          posX = pos[0].cast<float>();
          posY = pos[1].cast<float>();
        }

        string coord;
        if (!coordinates.is_none())
          coord = toLower((string)py::str(coordinates));
 
        float absX, absY;
        if (coord.empty() || strcmp(coord.c_str(), "top") == 0)
        {
          absX = float(caller->Left) + posX;
          absY = float(caller->Top) + posY;
        }
        else if (strcmp(coord.c_str(), "sheet") == 0)
        {
          absX = posX;
          absY = posY;
        }
        else if (strcmp(coord.c_str(), "bottom") == 0)
        {
          absX = float(caller->Left) + float(caller->Width) + posX;
          absY = float(caller->Top) + float(caller->Height) + posY;
        }
        else
          throw py::value_error("Coord argument is invalid");

        // Create temp file and call the file write function. Temporarily release
        // the GIL in case any file system issues cause a delay
        wstring tempFileName;
        {
          py::gil_scoped_release releaseGil;
          HANDLE tempFileHandle;
          std::tie(tempFileHandle, tempFileName) = Helpers::makeTempFile();
          CloseHandle(tempFileHandle);
        }

        checkUserException([&]() {saveFunction(tempFileName); });

        // Release GIL before doing any more slow stuff
        py::gil_scoped_release releaseGil;

        auto shapes = caller->Worksheet->Shapes;
        auto shapeName = wstring(L"XLOIMG_") + (const wchar_t*)(caller->GetAddressLocal(true, true, Excel::xlA1));

        // I don't think it's possible to check if the shape exists prior to deletion
        // so we have to catch the error

        // TODO: copy size info from existing image?
        try
        {
          shapes->Item(shapeName.c_str())->Delete();
        }
        catch (_com_error) {}

        auto newPic = shapes->AddPicture2(
          _bstr_t(tempFileName.c_str()),
          Office::MsoTriState::msoFalse,
          Office::MsoTriState::msoTrue,
          absX, absY,
          width, height,
          compress ? Office::msoPictureCompressTrue : Office::msoPictureCompressFalse);

        newPic->Name = shapeName.c_str();

        // Remove temporary file in a separate thread.
        std::thread([file = tempFileName]() {
          DeleteFile(file.c_str());
        }).detach();

        return shapeName;
      }
      XLO_RETHROW_COM_ERROR;
    }

    namespace
    {
      constexpr auto EXCEL_TLB_LCID  = 0;
      constexpr auto EXCEL_TLB_MAJOR = 1;
      constexpr auto EXCEL_TLB_MINOR = 4;

      py::object comObjectWithComtypes(
        IUnknown* p, const char* iface, const wchar_t* clsid)
      {
        // Important note: does not call QI on provided IUnknown, assumes it is a hanging reference
        auto comtypes = py::module::import("comtypes.client");

        auto libidVer = py::tuple(3);
        libidVer[0] = LIBID_STR_Excel;
        libidVer[1] = EXCEL_TLB_MAJOR;
        libidVer[2] = EXCEL_TLB_MINOR;
        auto typelib = comtypes.attr("GetModule")(libidVer);

        auto ptrType = typelib.attr(iface);

        auto ctypesPtr = py::module::import("ctypes").attr("POINTER")(ptrType)(
          PySteal<py::object>(PyLong_FromVoidPtr(p)));

        return comtypes.attr("_manage")(ctypesPtr, clsid, 1);
      }

      auto find_PyCom_PyObjectFromIUnknown()
      {
        // Do a GetProcAddress for 'PyCom_PyObjectFromIUnknown' in pythoncom.dll

        const std::wstring pythoncomDLL = py::module::import("pythoncom").attr("__file__").cast<std::wstring>();
        auto pythoncom = LoadLibrary(pythoncomDLL.c_str());
        if (!pythoncom)
          XLO_THROW(L"Failed to load pythoncom DLL '{}'", pythoncomDLL);
        
        auto funcAddress = GetProcAddress(pythoncom, "PyCom_PyObjectFromIUnknown");
        if (!funcAddress)
          XLO_THROW(L"Failed to find PyCom_PyObjectFromIUnknown in pythoncom DLL '{}'", pythoncomDLL);

        typedef PyObject* (*FuncType)(IUnknown*, REFIID riid, BOOL);
        return (FuncType)funcAddress;
      }

      py::object comObjectWithPyCom(
        IUnknown* p, const char* iface, const wchar_t* clsid)
      {
        // Calls QI on provided IUnknown
        static auto bindFunction = find_PyCom_PyObjectFromIUnknown();
        auto dispatchPtr = PySteal<py::object>(bindFunction(p, IID_IDispatch, true));

        auto targetType =
          py::module::import("win32com.client.gencache").attr("GetModuleForCLSID")(clsid).attr(iface);

        return targetType(dispatchPtr);
      }

      py::object marshalCom(
        const char* binderLib, IUnknown* p, const char* interfaceName, const GUID& clsid)
      {
        if (!binderLib || binderLib[0] == 0)
          return marshalCom(theCoreAddin().comBinder.c_str(), p, interfaceName, clsid);
        
        // Convert our CLSID to a string, 128 chars should be plenty
        wchar_t clsidStr[128];
        StringFromGUID2(clsid, clsidStr, _countof(clsidStr));

        if (_stricmp(binderLib, "comtypes") == 0)
        {
          // Must pass a appropriately QI'd hanging ref for comtypes
          IUnknown* q;
          if (p->QueryInterface(clsid, (void**)&q) != 0)
            XLO_THROW("Internal: failed to find COM interface '{}' for '{}'", interfaceName, binderLib);
          return comObjectWithComtypes(q, interfaceName, clsidStr);
        }
        else if (_stricmp(binderLib, "win32com") == 0)
          return comObjectWithPyCom(p, interfaceName, clsidStr);

        XLO_THROW("Unsupported COM lib '{}'", binderLib);
      }
    }

    py::object comToPy(Excel::_Application* p, const char* binder)
    {
      return marshalCom(binder, p, "_Application", __uuidof(Excel::_Application));
    }
    pybind11::object comToPy(Excel::Window* p, const char* binder)
    {
      return marshalCom(binder, p, "Window", __uuidof(Excel::Window));
    }
    pybind11::object comToPy(Excel::_Workbook* p, const char* binder)
    {
      return marshalCom(binder, p, "_Workbook", __uuidof(Excel::_Workbook));
    }
    pybind11::object comToPy(Excel::_Worksheet* p, const char* binder)
    {
      return marshalCom(binder, p, "_Worksheet", __uuidof(Excel::_Worksheet));
    }
    pybind11::object comToPy(Excel::Range* p, const char* binder)
    {
      return marshalCom(binder, p, "Range", __uuidof(Excel::Range));
    }

    namespace
    {
      py::object getExcelApp(const char* comlib)
      {
        return comToPy(&excelApp(), comlib);
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        mod.def("insert_cell_image", writeCellImage, 
          py::arg("writer"),
          py::arg("size") = py::none(),
          py::arg("pos") = py::none(),
          py::arg("origin") = py::none(),
          py::arg("compress") = true);

        mod.def("application", getExcelApp, py::arg("comlib")="");
      });
    }
  }
}
