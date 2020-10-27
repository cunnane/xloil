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

These are mostly documented under `Excel.Appliction <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_

The exceptions are  The `CalcCancelled` and `WorkbookAfterClose` Events
which are not part of the Application object.
       
.. highlight:: c++

::
     xloil::Event::CalcCancelled().bind(
          [self = this]() { self->cancel(); }));

          static auto handler = Event::AfterCalculate() += [logger]() { logger->flush(); };

