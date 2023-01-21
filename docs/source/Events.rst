==================
xlOil Excel Events
==================

.. contents::
    :local:

Introduction
------------

xlOil allows code to hook into Excel events.  The Excel API documentation 
`Excel.Application <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_
contains more complete descriptions of most of these events, with the exception of the
`CalcCancelled`, `WorkbookAfterClose`, `XllAdd` and `XllRemove` events which are not part of the
*Application* object and are provided by xlOil.

For the syntax used to hook these events, see the individual language documentation.

Parameter types
===============
Most events use some of the parameter types described below:

* *Workbook name*: passed as a string (not a ``Workbook`` object)
* *Worksheet name*: passed as a string 
* *Range*: passed as a ``Range`` object
* *Cancel*: a bool which starts as False If the event handler sets this argument to True,
  further processing of the event will stop
  
AfterCalculate
--------------

This event occurs after all Worksheet.Calculate, Chart.Calculate, QueryTable.AfterRefresh, 
and SheetChange events. It's the last event to occur after all refresh processing and all 
calc processing have completed, and it occurs after CalculationState is set to xlDone.
This event is called if the calculation was interrupted (for example by the user pressing 
a key or mouse button).

WorkbookOpen
------------
Occurs when a workbook is opened. Takes a single parameter containing the workbook name.

NewWorkbook
-----------
Occurs when a new workbook is created. Takes a single parameter containing the workbook name.

SheetSelectionChange
--------------------
Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a 
chart sheet). Takes two paramters:

* Worksheet name
* Target / selected Range

SheetBeforeDoubleClick
----------------------
Occurs when any worksheet is double-clicked, before the default double-click action. 
Takes three parameters

 * Worksheet name
 * Target - A Range giving cell nearest to the mouse pointer when the double-click occurred.
 * Cancel - False when the event occurs. If the event procedure sets this argument to True, the default 
   double-click action isn't performed when the procedure is finished.

SheetBeforeRightClick
---------------------
Occurs when any worksheet is right-clicked, before the default right-click action. 
Takes three parameters

 * Worksheet name
 * Target - A Range giving the cell nearest to the mouse pointer when the right-click occurred.
 * Cancel - False when the event occurs. If the event procedure sets this argument to True, the default 
   right-click action isn't performed when the procedure is finished.

SheetActivate
-------------
Occurs when any sheet is activated. Takes a single parameter containing the sheet name.

SheetDeactivate
---------------
Occurs when any sheet is deactivated. Takes a single parameter containing the sheet name.

SheetCalculate
--------------
Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.
Takes a single parameter containing the sheet name.

SheetChange
-----------
Occurs when cells in any worksheet are changed by the user or by an external link. Takes 
two paramters:

* Worksheet name
* Changed Range

WorkbookActivate
----------------
Occurs when any workbook is activated.  Takes a single parameter containing the workbook name.

WorkbookDeactivate
------------------
Occurs when any workbook is deactivated.  Takes a single parameter containing the workbook name.

WorkbookBeforeClose
-------------------
Occurs immediately before any open workbook closes. Takes two parameters

 * Workbook name
 * Cancel - False when the event occurs. If the event procedure sets this argument to True, the 
   workbook doesn't close when the procedure is finished.

The event is not called for each workbook when Excel exits.

WorkbookBeforeSave
------------------
Occurs before any open workbook is saved. Takes three parameters:

 * Workbook name
 * SaveAsUI - True if the Save As dialog box will be displayed due to changes made that need to 
   be saved in the workbook.
 * Cancel - False when the event occurs. If the event procedure sets this argument to True, the 
   workbook isn't save when the procedure is finished.

WorkbookBeforePrint
-------------------
Occurs immediately before any open workbook is printed. Takes two parameters

 * Workbook name
 * Cancel - False when the event occurs. If the event procedure sets this argument to True, the 
   workbook doesn't print when the procedure is finished.

WorkbookAfterClose
------------------
Excel's *WorkbookBeforeClose* event is cancellable by the user so it is not possible to know if 
the workbook actually closed.  When xlOil calls `WorkbookAfterClose`, the workbook is certainly 
closed, but it may be some time since that closure happened. Takes a single parameter containing the 
workbook name.

The event is not called for each workbook when xlOil exits. This event is not part of the 
*Excel.Application* API.

WorkbookNewSheet
----------------
Occurs when a new sheet is created in any open workbook. The first parameter is the workbook name, 
the second is the new sheet name.

WorkbookAddinInstall
--------------------
Occurs when a workbook is installed as an add-in. Takes a single parameter containing the workbook name.

WorkbookAddinUninstall
----------------------
Occurs when any add-in workbook is uninstalled. Takes a single parameter containing the workbook name.

CalcCancelled
-------------
Called when the calculation cycle is cancelled (for example by the user pressing a key or mouse button).
Native async functions should stop any background calculation when this event is received.

This event is not part of the *Excel.Application* API.

XllAdd
------
Triggered when an XLL related to this instance of xlOil is added by the user using the Addin settings
window. The parameter is the XLL filename.

This event is not part of the *Excel.Application* API.

XllRemove
---------
Triggered when an XLL related to this instance of xlOil is removed by the user using the Addin settings
window. The parameter is the XLL filename.

This event is not part of the *Excel.Application* API.
