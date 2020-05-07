#pragma once

#include <oleacc.h> // must include this before type imports

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
#import "libid:00020813-0000-0000-C000-000000000046"  \
  rename("DocumentProperties", "DocumentPropertiesXL") \
  rename( "DialogBox", "ExcelDialogBox" ) \
  rename( "RGB", "ExcelRGB" ) \
  rename( "CopyFile", "ExcelCopyFile" ) \
  rename( "ReplaceText", "ExcelReplaceText" )
