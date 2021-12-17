from . import _core

COM_LIB = "win32com"

def use_com_lib(name:str):
    """
        Selects the library used for COM support.  This impacts how the properties 
        and methods of the *Excel.Application* object returned by xloil.app() are 
        resolved.  THe choices are:

           * comtypes: a newer pure python package (which uses `ctypes`)
           * win32com: a well-established more C++ based library

        You must call this *before* any call to xloil.app()
    """
    global COM_LIB
    COM_LIB = name

_excel_application_obj = None

def _get_excel_application_win32com():
    import xloil_core
    import pythoncom
    import win32com.client

    punk = xloil_core.get_excel_app_pycom(pythoncom.__file__)
    return win32com.client.CastTo(punk, '_Application')


def _get_excel_application_comtypes():
    import xloil_core
    import comtypes
    import comtypes.client
    import ctypes

    clsid = comtypes.GUID.from_progid("Excel.Application")
    ptr = ctypes.POINTER(comtypes.IUnknown)(xloil_core.application())
    return comtypes.client._manage(ptr, clsid, None)

def app():
    """
        Returns a handle to the *Excel.Application* object.  This object is the root of 
        Excel's COM interface and supports a wide range of operations; it will be familiar 
        to VBA programmers.  It is well documented by Microsoft, see 
        https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview
        and https://docs.microsoft.com/en-us/office/vba/api/excel.application(object).
        
        Properties and methods of the `app()` object are resolved dynamically at runtime
        so cannot be seen by python inspection/completion tools.  xlOil uses a 3rd party
        library to do this resolution and handle all interation with the *Excel.Application*
        object.  This library defaults to `comtypes` but `win32com` can be choosen by
        calling `xloil.use_com_lib('win32com')` before any call to `xloil.app()`.

        Many operations using the Application object will only work in 
        functions declared as **macro type**.

        Examples
        --------

        To get the name of the active worksheet:

        ::
            
            @func(macro=True)
            def sheetName():
                return xlo.app().ActiveSheet.Name

    """
    global _excel_application_obj, COM_LIB

    if _excel_application_obj is None:
        choices = {
            "comtypes" : _get_excel_application_comtypes,
            "win32com" : _get_excel_application_win32com
        }
        try:
            _excel_application_obj = choices[COM_LIB]()
        except KeyError:
            raise KeyError(f"Unknown COM support library '{COM_LIB}', must be one of {choices.keys()}")

    return _excel_application_obj


class EventsPaused():
    """
    A context manager which stops Excel events from firing whilst
    the context is in scope
    """
    def __enter__(self):
        _core.event.pause()
        return self
    def __exit__(self, type, value, traceback):
        _core.event.allow()

    