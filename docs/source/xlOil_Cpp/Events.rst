================
xlOil C++ Events
================

xlOil allows code to hook into Excel events. The events supported are:

    *  AfterCalculate
    *  WorkbookOpen
    *  NewWorkbook
    *  SheetSelectionChange
    *  SheetBeforeDoubleClick
    *  SheetBeforeRightClick
    *  SheetActivate
    *  SheetDeactivate
    *  SheetCalculate
    *  SheetChange
    *  WorkbookAfterClose
    *  WorkbookActivate
    *  WorkbookDeactivate
    *  WorkbookBeforeClose
    *  WorkbookBeforeSave
    *  WorkbookBeforePrint
    *  WorkbookNewSheet
    *  WorkbookAddinInstall
    *  WorkbookAddinUninstall
    *  XllAdd
    *  XllRemove

These are mostly documented under `Excel.Appliction <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_

The exceptions are  The `CalcCancelled` and `WorkbookAfterClose` Events
which are not part of the Application object.
       
WorkbookAfterClose
------------------
Called after a workbook has closed (unlike WorkbookBeforeClose which can be cancelled by the user).

CalcCancelled
-------------
Called when the calculation cycle is cancelled (for example by the user pressing a key or mouse button).
Native async functions should stop any background calculation when this event is received.

XllAdd
------
Triggered when an XLL related to this instance of xlOil is added by the user using the Addin settings
window. The parameter is the XLL filename.

XllRemove
---------
Triggered when an XLL related to this instance of xlOil is removed by the user using the Addin settings
window. The parameter is the XLL filename.

Examples
--------

.. highlight:: c++

::

    // When the returned shared_ptr is destroyed, the handler is unhooked.
    auto ptr = xloil::Event::CalcCancelled().bind(
          [this]() { this->cancel(); }));

    // The returned id can be used to unhook the handler
    static auto id = Event::AfterCalculate() += [logger]() { logger->flush(); };

