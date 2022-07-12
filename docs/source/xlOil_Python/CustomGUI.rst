==============================
xlOil Python GUI Customisation
==============================

.. contents::
    :local:


Status Bar
----------

Possibly the simplest Excel GUI interaction: writing messages to Excel's status bar:

::

    from xloil import StatusBar

    with StatusBar(1000) as status:
        status.msg('Doing slow thing')
        ...
        status.msg('Done slow thing')

The `StatusBar` object clears the status bar after the specified number of milliseconds
once the `with` context ends


Ribbon
------

xlOil allows dynamic creation of Excel Ribbon components. See :any:`concepts-ribbon` for 
background.

::

    gui = xlo.create_gui(r'''<customUI xmlns=...>....</<customUI>''', 
        mapper={
            'onClick1': run_click,
            'onClick2': run_click,
            'onUpdate': update_text,
        })

The ``mapper`` dictionary (or function) links callbacks named in the ribbon XML to python functions. 
Each handler should have a signature like the following:

::

    def ribbon_callback1(ctrl: RibbonControl)
        ...
    def ribbon_callback2(ctrl: RibbonControl, arg1, arg2)
        ...
    def ribbon_callback3(ctrl: RibbonControl, *args)
        ...    
    async def ribbon_callback4(ctrl: RibbonControl, *args)
        ...    

The ``RibbonControl`` describes the control which raised the callback. The number of additional
arguments is callback dependent.  Some callbacks may be expected to return a value. 
See the *Resources* in :any:`concepts-ribbon` for a description of the appropriate callback signature.

Callbacks declared async will be executed in the addin's event loop. Other callbacks are executed 
in Excel's main thread. Async callbacks cannot return values.

The `getImage` callbacks must return a `PIL Image <https://pillow.readthedocs.io/en/stable/reference/Image.html>`_.
Instead of using a `getImage` per control, a single `loadImage` attribute can be added:

::

    <customUI loadImage="MyImageLoader" xmlns=...>
        ...
        <button id="AButton" image="icon.jpg" />

The `MyImageLoader` function will be called with the argument `icon.jpg` and be expected to return
a *PIL Image*.

Instead of a dictionary, the `mapper` object can be a function which takes any string and returns a 
callback handler.

The ``gui`` object returned above is actually a handle to a COM addin created to support
the ribbon customisation.  If the object goes out of scope and is deleted by python or if you call 
``ribbon.disconnect()``, the add-in is unloaded along with the ribbon customisation.

See :doc:`ExampleGUI` for an example of ribbon customisation.

Custom Task Panes
-----------------

`Custom task panes <https://docs.microsoft.com/en-us/visualstudio/vsto/custom-task-panes>`_ are user 
interface panels that are typically docked to one side of a window in Excel application.

Custom task panes are created using the `ExcelGUI` object. There is no need to create a ribbon as
well, but task panes are normally opened using a ribbon button.

Currently only Qt is supported using ``PyQt5`` or ``PySide2``. Additional support may be added.

::

    import xloil.gui.pyqt5                  # Must do this first!
    from PyQt5.QtWidgets import QWidget     # Could use PySide2 instead

    class MyTaskPane(QWidget):
        def __init__(self): # Must have no args
            ... # some code to draw the widget
        def send_signal(int):
            ... # some code to emit a Qt signal

    excelui = xlo.create_gui(...)
    pane = excelui.create_task_pane('MyPane', creator=MyTaskPane)

    pane.widget.send_signal(3)

The :any:`xloil.create_task_pane` call first looks for a pane with the specified name which is already 
attached to the active window, returning a reference to it if found.  Otherwise the ``creator``
is used.  If ``creator`` inherits from `QWidget`, it is constructed and attached to a new
custom task pane

It is also possible to pass a function as the ``creator`` argument.  The function should take an
:obj:`xloil.TaskPaneFrame` and return a :obj:`xloil.CustomTaskPane`.

To talk to your widget, it's best to set up a system of Qt 
`signals <https://wiki.qt.io/Qt_for_Python_Signals_and_Slots>`_. 
(the `syntax differs slightly in PyQt5 <https://www.pythonguis.com/faq/pyqt5-vs-pyside2/>`_).


Qt Thread-safety
================

With Qt, all GUI interactions (other than signals) must take place in the same thread, or 
Qt will abort.  To achieve this, xlOil creates a special Qt thread running the Qt event loop, 
then constructs ``MyTaskPane`` on that thread.

To run GUI commands on xlOil's *Qt* thread, do the following:

::

    from xloil.gui.pyqt5 import Qt_thread
    future = Qt_thread().submit(func, args) # Qt_thread is a concurrent.futures.Executor
    future.result()                         # Optional if result is required now



Task Pane registry
==================

The ``pane`` object is automatically stored in a registry so there is no need to hold a reference.
Task panes are attached by default to the active window and it is possible to have multiple 
windows per open workbook.  xlOil will free the panes when the parent workbook closes.

To look in the regitry for a task pane without having a :obj:`xloil.ExcelGUI` object:

::

    pane = xloil.find_task_pane("MyPane")





