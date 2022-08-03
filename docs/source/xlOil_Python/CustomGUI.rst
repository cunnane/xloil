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
interface panels that are usually docked to one side of a window in Excel application.

Custom task panes are created using the `ExcelGUI` object. There is no need to create a ribbon as
well, but task panes are normally opened using a ribbon button.


Qt Custom Task Panes
====================

Qt support uses *PyQt5* or *PySide2* (but not both!). The examples below use *PyQt5* but
*PySide2* can be substituted in place.  

.. caution::
    You *must* import :any:`xloil.gui.pyqt5` (or `xloil.gui.pyside2`) before any other
    use of Qt.  This allows xlOil to create and the *QApplication* on its own thread.

It's common in Qt GUIs to inherit from `QWidget`, so xlOil allows you to create a pane
from a `QWidget`:

::

    import xloil.gui.pyqt5  # or import xloil.gui.pyside2       
    from PyQt5.QtWidgets import QWidget     

    class QtTaskPane(QWidget):
        def __init__(self):
            super().__init__() # Don't forget this!
            ... # some code to draw the widget
        def send_signal(self, int):
            ... # some code to emit a Qt signal

    excelui = xlo.create_gui(...)
    pane = excelui.attach_pane('MyPane', pane=QtTaskPane)

    # The widget is in the pane's `widget` attribute
    pane.widget.send_signal(3) 

The :any:`xloil.ExcelGUI.attach_pane` call creates a task pane with the specified name.  If ``pane`` 
is a *type* which inherits from `QWidget`, it is constructed (on the Qt thread, see below)
and placed in a :any:`xloil.gui.pyqt5.QtThreadTaskPane` then attached to the Excel window.

To talk to your widget, it's best to set up a system of Qt 
`signals <https://wiki.qt.io/Qt_for_Python_Signals_and_Slots>`_ as these are thread-safe. 
(Note the `syntax differs slightly in PyQt5 <https://www.pythonguis.com/faq/pyqt5-vs-pyside2/>`_) 


Qt Thread-safety
________________

All *Qt* interactions other than signals must take place in the same thread, or Qt
will abort.  xlOil creates a special Qt thread which runs the Qt event loop, and 
constructs any task panes on that thread.

To run commands on xlOil's *Qt* thread, use the :any:`xloil.gui.pyqt5.Qt_thread` object

::

    from xloil.gui.pyqt5 import Qt_thread
    future = Qt_thread().submit(func, *args)        # returns a concurrent.futures.Future
    future2 = Qt_thread().submit_async(func, *args) # returns an asyncio.Future
    future.result()                                 # Blocks if result is required now

You can also use `Qt_thread` as a decorator to wrap the function in a `submit` call, for example:

::

    @Qt_thread
    def some_func():
        ...

Tkinter Custom Task Panes
=========================

We create a class which derives from :any:`xloil.gui.tkinter.TkThreadTaskPane` (which in turn 
derives from :any:`xloil.gui.CustomTaskPane`).  Unlike Qt, it's not common to derive the from 
a *tkinter* object.

We draw the window into the *tkinter.Toplevel* contained in `self.top_level`.

::
    
    from xloil.gui.tkinter import TkThreadTaskPane, Tk_thread

    class TkTaskPane(TkThreadTaskPane):
    
        @Tk_thread
        def set_x(self, x):
            ...
        
        def __init__(self):
            super().__init__() # Don't forget this!
            
            # This name is picked up by ExcelGUI.attach_pane
            self.name = "MyPane"

            import tkinter as tk
            
            top_level = self.top_level
            # Draw into window
            ...
            

    excelui = xlo.ExcelGUI(xml=..., funcmap=...)
    pane = excelui.attach_pane(TkTaskPane())

    pane.set_x(3)

As *tkinter* does not have thread-safe signals, we use must ensure `set_x` is run on the correct
thread. The :any:`xloil.gui.tkinter.Tk_thread` function behaves the same as `Qt_thread` described
in :ref:`xlOil_Python/CustomGUI:Qt Thread-safety`.  The `__init__` method is always called on the 
*tkinter* thread so we don't need to decorate it.


Task Pane registry
==================

The created task panes are automatically stored in a registry so there is no need to hold a
reference to them.  Task panes are attached by default to the active window and it is possible to 
have multiple windows per open workbook.  xlOil will free the panes when the parent workbook ora
addin closes.

We can search the registry by name for a task pane without having the :obj:`xloil.ExcelGUI` object:

::

    pane = xloil.find_task_pane("MyPane")

By default, xlOil looks for a pane attached to the active window, but this can be changed witha
arguments.  It is possible to create multiple panes with the same name, in which case this search
could return either one.


Async GUI Calls
---------------

The above examples create the GUI calls in a synchronous fashion but many of the GUI functions
are or can be async.  Because xlOil loads modules in a background thread, it's not necessary
to do this to keep Excel responsive but it could be useful in some circumstances.

::

    async def make_gui():
    
        # With connect=False the ctor does nothing
        excelui = xlo.ExcelGUI(xml=..., funcmap=..., connect=False)

        # The action happens when we call connect, which returns a awaitable future
        await excelui.connect()

        # We can also create the pane async by passing an awaitable but we have 
        # to then pass the name explictly
        await excelui.attach_pane_async(
            name='TkPane',
            pane=Tk_thread().submit_async(TkTaskPane))

        # We need to keep a reference to 'excelui' as deleting it disconnects the UI
        return excelui, pane
