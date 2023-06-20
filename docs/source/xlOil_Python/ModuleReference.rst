=================================
xlOil Python Module Reference
=================================

.. contents::
    :local:

Declaring Worksheet Functions
-----------------------------

.. currentmodule:: xloil

.. autosummary::
	AllowRange
	Arg
	Array
	Cache
	CannotConvert
	CellError
	ExcelArray
	FastArray	
	SingleValue
	func
	converter
	returner
	import_functions
	register_functions
	deregister_functions
	scan_module

.. automodule:: xloil
	:members: Arg,Array,Cache,CannotConvert,CellError,ExcelArray,FastArray,SingleValue,func,converter,returner,import_functions,register_functions,deregister_functions,scan_module
	:imported-members:
	:undoc-members:

.. autodata:: AllowRange

Excel Object Model
------------------

.. currentmodule:: xloil

.. autosummary::
	workbooks
	worksheets
	app
	active_worksheet
	active_workbook
	Application
	Caller
	Range
	Workbook
	Worksheet
	ExcelWindow
	ExcelWindows
	Workbooks
	Worksheets

.. autodata:: workbooks
	
.. autodata:: worksheets
	
.. autofunction:: app

.. autofunction:: active_worksheet

.. autofunction:: active_workbook

.. autoclass:: Application
	:members: 
	:inherited-members:
	:undoc-members:
	:special-members: __enter__, __exit__

.. autoclass:: Caller
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: Range
	:members: 
	:inherited-members:
	:undoc-members:
	:special-members: __getitem__

.. autoclass:: Workbook
	:members: 
	:inherited-members:
	:undoc-members:
	:special-members: __getitem__

.. autoclass:: Worksheet
	:members: 
	:inherited-members:
	:undoc-members:
	:special-members: __getitem__

.. autoclass:: ExcelWindow
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: ExcelWindows
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: Workbooks
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: Worksheets
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: PauseExcel
	:members: 

RTD Functions
-------------

.. currentmodule:: xloil

.. autosummary::
	RtdPublisher
	RtdServer
	xloil.rtd.subscribe
	xloil.rtd.RtdSimplePublisher

.. autoclass:: RtdPublisher
	:members:
	
.. autoclass:: RtdServer
	:members:

.. automodule:: xloil.rtd
	:members:


GUI Interaction
---------------

.. currentmodule:: xloil

.. autosummary::
	StatusBar
	ExcelGUI
	TaskPaneFrame
	RibbonControl
	xloil.gui.CustomTaskPane
	xloil.gui.find_task_pane
	xloil.gui.qtpy.Qt_thread
	xloil.gui.qtpy.QtThreadTaskPane
	xloil.gui.tkinter.Tk_thread
	xloil.gui.tkinter.TkThreadTaskPane
	xloil.gui.wx.wx_thread
	xloil.gui.wx.WxThreadTaskPane

.. autoclass:: StatusBar
	:members:

.. autoclass:: ExcelGUI
	:members:
.. autoclass:: TaskPaneFrame
	:members:
.. autoclass:: RibbonControl
	:members:

.. automodule:: xloil.gui
	:members: CustomTaskPane 

.. autofunction:: find_task_pane

.. automodule:: xloil.gui.qtpy
	:members: Qt_thread, QtThreadTaskPane
	:inherited-members:

.. automodule:: xloil.gui.tkinter
	:members: Tk_thread, TkThreadTaskPane	
	:inherited-members:

.. automodule:: xloil.gui.wx
	:members: wx_thread, WxThreadTaskPane	
	:inherited-members:

Events
------

.. currentmodule:: xloil.event

.. automodule:: xloil.event

.. autoclass:: Event
	:members:
	:special-members: __iadd__, __isub__

.. autofunction:: pause
.. autofunction:: allow

.. autodata:: AfterCalculate
.. autodata:: WorkbookOpen
.. autodata:: NewWorkbook
.. autodata:: SheetSelectionChange
.. autodata:: SheetBeforeDoubleClick
.. autodata:: SheetBeforeRightClick
.. autodata:: SheetActivate
.. autodata:: SheetDeactivate
.. autodata:: SheetCalculate
.. autodata:: SheetChange
.. autodata:: WorkbookAfterClose
.. autodata:: WorkbookRename
.. autodata:: WorkbookActivate
.. autodata:: WorkbookDeactivate
.. autodata:: WorkbookBeforeClose
.. autodata:: WorkbookBeforeSave
.. autodata:: WorkbookAfterSave
.. autodata:: WorkbookBeforePrint
.. autodata:: WorkbookNewSheet
.. autodata:: WorkbookAddinInstall
.. autodata:: WorkbookAddinUninstall
.. autodata:: XllAdd
.. autodata:: XllRemove
.. autodata:: ComAddinsUpdate
.. autodata:: PyBye
.. autodata:: UserException


Everything else
---------------

.. currentmodule:: xloil

.. autosummary::
	in_wizard
	get_async_loop
	get_event_loop
	from_excel_date
	date_formats
	linked_workbook
	source_addin
	excel_state
	run
	run_async
	call
	call_async
	excel_callback
	cache
	Addin
	source_addin
	xloil_addins
	core_addin
	xloil._core._AddinsDict
	xloil._core._DateFormatList
	xloil._core._LogWriter
	xloil.logging.log
	xloil.debug.use_debugger

.. autoclass:: ObjectCache
	:members: 

.. autodata:: cache
	:annotation: = ObjectCache
	:no-value:

.. automodule:: xloil
	:members: in_wizard,get_async_loop,get_event_loop,from_excel_date,linked_workbook,source_addin,excel_state,run,run_async,call,call_async,excel_callback,source_addin,xloil_addins,core_addin
	:imported-members:
	:undoc-members:

.. autodata:: date_formats
	:annotation: = _DateFormatList
	:no-value:

.. autoclass:: ExcelState
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: Addin
	:members:

.. autodata:: xloil_addins
	:annotation: = _AddinsDict
	:no-value:

.. automodule:: xloil._core
	:members: _AddinsDict, _DateFormatList, _LogWriter

.. automodule:: xloil.logging
	:members: 

.. automodule:: xloil.debug
	:members:



External libraries
------------------

.. currentmodule:: xloil

.. autosummary::
	insert_cell_image
	xloil.pandas.PDFrame
	xloil.pillow.ReturnImage
	xloil.matplotlib.ReturnFigure

.. autofunction:: insert_cell_image

.. automodule:: xloil.pandas
	:members: PDFrame
.. automodule:: xloil.pillow
	:members: ReturnImage
.. automodule:: xloil.matplotlib
	:members: ReturnFigure

