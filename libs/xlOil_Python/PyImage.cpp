#include "PyImage.h"
#include "PyHelpers.h"
#include <xlOil/State.h>

#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <olectl.h>

namespace py = pybind11;
namespace xloil
{
  namespace Python
  {
    namespace
    {
      // Ensure handle is cleaned on exception
      struct HandleDC
      {
        HandleDC(HDC h_) : h(h_) {}
        ~HandleDC() { DeleteDC(h); }
        HDC h;
      };
      struct HandleBitmap
      {
        HandleBitmap(HBITMAP h_) : h(h_) {}
        ~HandleBitmap() { DeleteObject(h); }
        void release() { h = nullptr; }
        HBITMAP h;
      };
    }
    IPictureDisp* pictureFromPilImage(const py::object& image)
    {
      // If PIL hasn't been imported, it's definitely not a PIL image!
      const auto sysModules = PyBorrow<py::dict>(PySys_GetObject("modules"));
      if (!sysModules.contains("PIL"))
        return nullptr;

      // Import PIL, and check if object is an instance of a PIL Image
      auto PIL = py::module::import("PIL.Image");
      auto pilImageType = PIL.attr("Image");
      if (!py::isinstance(image, pilImageType))
        return nullptr;
      
      // We have an Image, so get its size
      const auto size = py::cast<py::tuple>(image.attr("size"));
      const auto width = py::cast<int>(size[0]);
      const auto height = py::cast<int>(size[1]);

      // Retrieve the handle to a display device context for the Excel window 
      const auto hwnd = (HWND)Environment::excelProcess().hWnd;
      auto hdcWindow = GetDC(hwnd);

      // Create a colour compatible device context and bitmap to receive the image
      auto memoryDC = HandleDC(CreateCompatibleDC(hdcWindow));
      auto bitmap = HandleBitmap(CreateCompatibleBitmap(hdcWindow, width, height));

      // Select our new bitmap in the device context
      SelectObject(memoryDC.h, bitmap.h);
      
      // No longer need the Excel window DC
      ReleaseDC(hwnd, hdcWindow);

      // Create a Device Independent Bitmap in PIL and copy it to the device context
      auto ImageWin = py::module::import("PIL.ImageWin");
      auto dib = ImageWin.attr("Dib")(image);
      dib.attr("expose")((intptr_t)memoryDC.h);
     
      // Get the (updated) palette from our device context
      auto palette = (HPALETTE)GetCurrentObject(memoryDC.h, OBJ_PAL);

      //SelectObject(memoryDC, oldbmp); Doesn't seem to be required
      
      // Call OleCreatePictureIndirect to create a IPictureDisp
      PICTDESC pd;
      pd.picType        = PICTYPE_BITMAP;
      pd.cbSizeofstruct = sizeof(PICTDESC);
      pd.bmp.hbitmap    = bitmap.h;
      pd.bmp.hpal       = palette;
      void* result;
      OleCreatePictureIndirect(&pd, __uuidof(IPictureDisp), TRUE, &result);

      // Release bitmap handle
      bitmap.release();

      return (IPictureDisp*)result;
    }
  }
}