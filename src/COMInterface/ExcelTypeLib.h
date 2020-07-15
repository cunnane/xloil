#pragma once

#include <oleacc.h> // must include this before type imports
#include <comdef.h>

// MSO.dll
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" \
  rename("DocumentProperties", "MSODocumentProperties") \
  rename("RGB", "MSORGB")

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


#define XLO_RETHROW_COM_ERROR \
  catch (_com_error& error)\
  {\
    XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());\
  }