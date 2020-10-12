#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <oleacc.h> // must include this before type imports
#include <comdef.h>

// MSO.dll
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" \
  rename("DocumentProperties", "MSODocumentProperties") \
  rename("RGB", "MSORGB") \
  rename("IAccessible", "MSOIAccessible")

using namespace Office;

// VBE6EXT.OLB
#import "libid:0002E157-0000-0000-C000-000000000046" \
  rename("Reference", "ignorethis")

using namespace VBIDE;

// Excel.exe
#import "libid:00020813-0000-0000-C000-000000000046" \
  rename("DocumentProperties", "DocumentPropertiesXL") \
  rename("DialogBox", "ExcelDialogBox" ) \
  rename("RGB", "ExcelRGB" ) \
  rename("CopyFile", "ExcelCopyFile" ) \
  rename("ReplaceText", "ExcelReplaceText" )

extern "C" const GUID __declspec(selectany) LIBID_Excel =
  { 0x00020813,0x0000,0x0000,{0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46} };

extern "C" const GUID __declspec(selectany) LIBID_AddInDesigner =
  { 0xAC0714F2,0x3D04,0x11D1, {0xAE,0x7D,0x00,0xA0,0xC9,0x0F,0x26,0xF4} };

//#import "C:\\Program Files\\Common Files\\Designer\\MSADDNDR.OLB"
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"

#define XLO_RETHROW_COM_ERROR \
  catch (_com_error& error)\
  {\
    XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());\
  }