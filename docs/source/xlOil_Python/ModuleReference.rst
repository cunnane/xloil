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
	Event
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
	LogWriter
	get_async_loop
	get_event_loop
	from_excel_date
	register_functions
	deregister_functions
	linked_workbook
	source_addin
	excel_callback
	excel_state
	run
	run_async
	call
	call_async	
	app
	Range
	Worksheet
	Workbook
	ExcelWindow
	windows
	workbooks
	active_worksheet
	active_workbook
	ExcelGUI
	CustomTaskPane
	TaskPaneFrame
	find_task_pane
    create_task_pane
	insert_cell_image
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

.. autodata:: windows

.. autodata:: workbooks
	
.. autofunction:: app

.. autofunction:: active_worksheet

.. autofunction:: active_workbook

.. autoclass:: Range
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: Workbook
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: Worksheet
	:members: 
	:inherited-members:
	:undoc-members:

.. autoclass:: ExcelWindow
	:members: 
	:inherited-members:
	:undoc-members:


RTD Functions
-------------

.. currentmodule:: xloil

.. autoclass:: RtdPublisher

.. autoclass:: RtdServer

.. automodule:: xloil.rtd
	:members:


GUI Interaction
---------------

.. currentmodule:: xloil

.. autoclass:: ExcelGUI
.. autoclass:: CustomTaskPane 
.. autoclass:: TaskPaneFrame
.. autoclass:: RibbonControl
.. autofunction:: find_task_pane
.. autofunction:: create_task_pane

.. automodule:: xloil.qtgui
	:members: Qt_thread, QtThreadTaskPane	


Everything else
---------------

.. automodule:: xloil
	:members: Event,in_wizard,LogWriter,get_async_loop,get_event_loop,from_excel_date,linked_workbook,source_addin,excel_run,excel_state,run,run_async,call,call_async
	:imported-members:
	:undoc-members:

.. autodata:: log

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

