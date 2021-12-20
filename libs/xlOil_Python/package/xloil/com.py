from . import _core

_excel_application_obj = None

def app(lib=None):
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
                return xlo.app().ActiveWorksheet.Name

    """
    global _excel_application_obj

    if lib is not None:
        import xloil_core
        return xloil_core.application(lib)

    if _excel_application_obj is None:
        import xloil_core
        _excel_application_obj = xloil_core.application()
       
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

    