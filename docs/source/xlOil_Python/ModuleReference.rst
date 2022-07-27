=================================
xlOil Python Module Reference
=================================

.. contents::
    :local:

.. currentmodule:: xloil

.. autosummary::
	Arg
	CannotConvert
	CellError
	ExcelArray
	AllowRange
	RtdPublisher
	RtdServer
	Cache
	SingleValue
	RibbonControl
	func
	converter
	returner
	in_wizard
	log
	get_async_loop
	get_event_loop
	from_excel_date
	register_functions
	deregister_functions
	linked_workbook
	source_addin
	excel_callback
	excel_state
	ExcelState
	run
	run_async
	call
	call_async	
	Caller
	Application
	Range
	Worksheet
	Workbook
	ExcelWindow
	Workbooks
	Worksheets
	app
	workbooks
	active_worksheet
	active_workbook
	ExcelGUI
	xloil.gui.CustomTaskPane
	TaskPaneFrame
	xloil.gui.find_task_pane
	insert_cell_image
	xloil.rtd.subscribe
	xloil.rtd.RtdSimplePublisher
	xloil.debug.exception_debug
	xloil.qtgui.Qt_thread
	xloil.pandas.PDFrame
	xloil.pillow.ReturnImage
	xloil.matplotlib.ReturnFigure
..
	[comment]: need to patch is_filtered_inherited_member in autodoc/__init__.py to get 
	[comment]: inherited-members to work, then can remove this horrible explict list


Declaring Worksheet Functions
-----------------------------

.. automodule:: xloil
	:members: Arg,CannotConvert,CellError,ExcelArray,Cache,SingleValue,func,converter,returner,register_functions,deregister_functions
	:imported-members:
	:undoc-members:

.. autodata:: AllowRange

Excel Object Model
------------------

.. currentmodule:: xloil

.. autodata:: workbooks
	
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

RTD Functions
-------------

.. currentmodule:: xloil

.. autoclass:: RtdPublisher
	:members:
	
.. autoclass:: RtdServer
	:members:

.. automodule:: xloil.rtd
	:members:


GUI Interaction
---------------

.. currentmodule:: xloil

.. autoclass:: ExcelGUI
	:members:
.. autoclass:: TaskPaneFrame
	:members:
.. autoclass:: RibbonControl
	:members:

.. automodule:: xloil.gui
	:members: CustomTaskPane 

.. autofunction:: find_task_pane

.. automodule:: xloil.gui.pyqt5
	:members: Qt_thread, QtThreadTaskPane
	:inherited-members:

.. automodule:: xloil.gui.tkinter
	:members: Tk_thread, TkThreadTaskPane	
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

.. automodule:: xloil
	:members: in_wizard,get_async_loop,get_event_loop,from_excel_date,linked_workbook,source_addin,excel_state,run,run_async,call,call_async,excel_callback
	:imported-members:
	:undoc-members:

.. autoclass::ExcelState
	:members:
	:imported-members:
	:undoc-members:
	:inherited-members:
	:private-members:
	:special-members:
	
.. autodata:: log

.. automodule:: xloil.logging
	:members: _LogWriter

.. automodule:: xloil.debug
	:members:

External libraries
------------------

.. currentmodule:: xloil

.. autofunction:: insert_cell_image

.. automodule:: xloil.pandas
	:members: PDFrame
.. automodule:: xloil.pillow
	:members: ReturnImage
.. automodule:: xloil.matplotlib
	:members: ReturnFigure

