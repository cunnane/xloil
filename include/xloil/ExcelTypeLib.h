/// <summary>
/// This header file imports the Excel typelib required for accessing the
/// COM interface, e.g. through the Excel::Application object. You may need
/// to disable multi-threaded compliation in any source files which include 
/// this
/// </summary>

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
  rename("IAccessible", "MSOIAccessible") \
  rename("SearchPath", "MSOSearchPath")

// VBE6EXT.OLB
#import "libid:0002E157-0000-0000-C000-000000000046" \
  rename("Reference", "VBReference")

//disable: automatically excluding 'name' while importing type library 'library'
#pragma warning(disable : 4192) 

// Excel.exe
#import "libid:00020813-0000-0000-C000-000000000046" \
  rename("DocumentProperties", "ExcelDocumentProperties") \
  rename("DialogBox", "ExcelDialogBox" ) \
  rename("RGB", "ExcelRGB" ) \
  rename("CopyFile", "ExcelCopyFile" ) \
  rename("ReplaceText", "ExcelReplaceText" )

extern "C" const GUID __declspec(selectany) LIBID_Excel =
  { 0x00020813,0x0000,0x0000,{0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46} };

extern "C" const GUID __declspec(selectany) LIBID_AddInDesigner =
  { 0xAC0714F2,0x3D04,0x11D1, {0xAE,0x7D,0x00,0xA0,0xC9,0x0F,0x26,0xF4} };

// C:\\Program Files\\Common Files\\Designer\\MSADDNDR.OLB
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"


/// <summary>
/// Code for VBA Ignore error generated from Excel COM.
/// 
/// From MSDN:
/// Excel will return this when you try to invoke the object model 
/// when the property browser is suspended.  This will happen around 
/// user edits to ensure that things don't get out of whack with 
/// automation slipping in in the middle.  The problem with VBA_E_IGNORE 
/// is that it is non-standard.  The COM prescribed way of addressing this
/// issue is to register an IMessageFilter implementation.  This allows 
/// COM to notify the server whenever another thread is trying to make a 
/// call and this gives the server the opportunity to reject the call if 
/// they aren't in a position to handle it. However, VBA_E_IGNORE occurs 
/// outside of this mechanism so you will have to roll your own mechanism 
/// of handling it.  I would suggest that you create a some sort of dispatch
/// loop whereby you handle the specific COM exception by waiting a few 
/// seconds and then jumping back to the start of the loop.
/// See https://social.msdn.microsoft.com/Forums/vstudio/en-US/9168f9f2-e5bc-4535-8d7d-4e374ab8ff09/hresult-800ac472-from-set-operations-in-excel
/// </summary>
constexpr HRESULT VBA_E_IGNORE = 0x800ac472;

/// <summary>
/// Catches COM errors (which do not derive from std::execption). This 
/// macro is automatically invoked just before XLO_FUNC_END in functions 
/// which access the COM interface (e.g. via Excel::Application).
/// </summary>
#define XLO_RETURN_COM_ERROR \
  catch (_com_error& error) \
  { \
    return xloil::returnValue(xloil::formatStr(L"#COM ERROR %X: %s", (size_t)error.Error(), error.ErrorMessage())); \
  }

/// <summary>
/// Catches COM errors (which do not derive from std::execption) and 
/// re-throws them as an xloil::Exception. In addition, it detects the
/// dreaded VBA_E_IGNORE and throws an xloil::ComBusyException in this
/// case.
/// </summary>
#define XLO_RETHROW_COM_ERROR \
  catch (_com_error& error) \
  { \
    if (error.Error() == VBA_E_IGNORE) \
      throw xloil::ComBusyException(); \
    else \
      XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage()); \
  }
  
