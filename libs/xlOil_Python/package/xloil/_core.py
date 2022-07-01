from ._common import *

import xloil_core

# TODO: how about from xloil_core import *?
from xloil_core import (
    CellError, ExcelArray, in_wizard, 
    event, cache, RtdServer, RtdPublisher,
    deregister_functions, get_async_loop,
    ExcelGUI, create_gui, 
    excel_callback, excel_state, ExcelState,
    Caller,
    CannotConvert, 
    from_excel_date,
    insert_cell_image,
    TaskPaneFrame,
    RibbonControl,
    StatusBar,
    app,
    Application, Range, ExcelWindow, Workbook, Worksheet, ExcelWindows, Workbooks, Worksheets,
    active_worksheet, active_workbook,
    run, run_async, call, call_async
)

#
# If we are being called from an xlOil embedded interpreter, we can import
# the symbols directly. Otherwise we define skeletons of the imported 
# types to support type-checking, linting, auto-completion and documentation.
#
if XLOIL_EMBEDDED:
    from xloil_core import workbooks
else:
    workbooks:Workbooks = None
    """
        Collection of all open workbooks as Workbook objects.
    
        Examples
        --------

            workbooks['MyBook'].path
            windows.active.workbook.path

    """
