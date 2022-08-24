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

The :any:`xloil.StatusBar` object clears the status bar after the specified number of milliseconds
once the `with` context ends


Ribbon
------

xlOil allows dynamic creation of Excel Fluent Ribbon components. See :any:`concepts-ribbon` for 
background.

::

    gui = xlo.ExcelGUI(r'''<customUI xmlns=...>....</<customUI>''', 
        funcmap={
            'onClick1': run_click,
            'onClick2': run_click,
            'onUpdate': update_text,
        })

The ``gui`` object :any:`xloil.ExcelGUI` holds a handle to a COM addin created to support
the ribbon customisation.  If the object goes out of scope and is deleted by python or if you call 
``gui.disconnect()``, the add-in is unloaded along with the ribbon customisation and any custom task 
panes.

The ``funcmap`` dictionary (or function) links callbacks named in the ribbon XML to python functions. 
Each handler should have a signature like the following:

::

    def ribbon_button_press(ctrl: RibbonControl)
        ...
    def ribbon_callback2(ctrl: RibbonControl, arg1, arg2)
        ...
    def ribbon_get_text_label(ctrl: RibbonControl, *args)
        return "something"
    async def ribbon_callback4(ctrl: RibbonControl, *args)
        ...    

The ``ctrl`` argument points to the control which raised the callback. The number of additional
arguments is callback dependent.  Some callbacks are expected to return a value. 
See `Customizing the Office Fluent Ribbon <https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338199(v=office.12)>`_
for a description of the appropriate callback signature.

.. note::

    Callbacks are executed in Excel's main thread unless declared *async*, in which case they will be 
    executed in the addin's event loop.  Async callbacks cannot return values.


Instead of a dictionary, the `funcmap` object can be a function which takes any string and returns a 
callback handler.

See :doc:`ExampleGUI` for an example of ribbon customisation.

Setting button images
=====================

Any `getImage` callbacks must return a `PIL Image <https://pillow.readthedocs.io/en/stable/reference/Image.html>`_.
This is converted to an appropriate COM object by xlOil. Instead of using a separate `getImage` callback 
per control, a single `loadImage` attribute can be added:

::

    <customUI loadImage="MyImageLoader" xmlns=...>
        ...
        <button id="AButton" image="icon.jpg" />

The `MyImageLoader` function will be called with the argument `icon.jpg` and be expected to return
a *PIL Image*.


Custom Task Panes
-----------------

`Custom task panes <https://docs.microsoft.com/en-us/visualstudio/vsto/custom-task-panes>`_ are user 
interface panels that are usually docked to one side of a window in the Excel application. They can 
contain a *Qt*, *Tk* or *wx* interface, or any suitable custom COM control. 

Custom task panes are created using the :any:`xloil.ExcelGUI` object. There is no need to create a ribbon 
as well, but task panes are normally opened using a ribbon button, because Excel does not provide a 
default way for users to show or hide custom task panes.

Custom task panes are associated with a document frame window, which presents a view of a workbook 
to the user.  If you want to display a custom task pane with multiple workbooks, create a new instance 
of the custom task pane when the user creates or opens a workbook. To do this, either handle the 
`WorkbookOpen` event, or require the user to press a ribbon button to open a task pane for the active
workbook.

Thread-safety
=============

The :any:`xloil.ExcelGUI` object and custom task panes can be created in any thread (internally they 
re-direct calls to Excel's main thread). Typically GUI creation will be done on xlOil's python loader 
thread, which also contains an *asyncio* event loop. The individual GUI toolkits are generally not 
thread-safe and should only be accessed from dedicated threads which xlOil creates.  This is described 
below per toolkit in more detail.

.. caution:

    If another non-xlOil Excel addin uses the same GUI toolkit, it is very likely that Excel will crash.

Qt Custom Task Panes
====================

Qt support uses *qtpy* which auto-detects the Qt bindings (PySide/PyQt) and standardises the 
small syntactic differences between the libraries.

.. caution::
    You *must* import :any:`xloil.gui.qtpy` before any other use of Qt.  This allows xlOil 
    to create and the *QApplication* on its own thread.

It's common in Qt GUIs to inherit from `QWidget`, so xlOil allows you to create a pane
from a `QWidget`:

::

    import xloil.gui.qtpy
    from qtpy.QtWidgets import QWidget     

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
and placed in a :any:`xloil.gui.qtpy.QtThreadTaskPane` then attached to the Excel window.

To talk to your widget, it's best to set up a system of Qt 
`signals <https://wiki.qt.io/Qt_for_Python_Signals_and_Slots>`_ as these are thread-safe. 
(Note the `syntax differs slightly in PyQt5 <https://www.pythonguis.com/faq/pyqt5-vs-pyside2/>`_
but it is standardised by *qtpy*) 


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
derives from :any:`xloil.gui.CustomTaskPane`).  Unlike Qt, it's not (I think) as common to derive
from a *tkinter.Frame* object.

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

As *tkinter* does not have thread-safe signals, although it does have events which could be used here, 
but for illustration, we ensure `set_x` is run on the *Tk* thread, by decorating it with 
:any:`xloil.gui.tkinter.Tk_thread`.  The `__init__` method is always called on the *tkinter* thread 
so we don't need to decorate it.

Tkinter Thread-safety
_____________________

The :any:`xloil.gui.tkinter.Tk_thread` function behaves the same as `Qt_thread` described
in :ref:`xlOil_Python/CustomGUI:Qt Thread-safety`. 


wxPython Custom Task Panes
==========================

It's common in wx GUIs to inherit from `wx.Frame`, so xlOil allows you to create a pane
from a `wx.Frame`:

::

    from xloil.gui.wx import wx_thread
    import wx

    class OurWxPane(wx.Frame):
        def __init__(self):
            super().__init__(None, title='Hello')
            ...

        @wx_thread
        def set_progress(self, x: int):
            ...

    excelui = xlo.create_gui(...)
    pane = excelui.attach_pane('MyPane', pane=OurWxPane)

    # The frame is in the pane's `frame` attribute
    pane.frame.set_progress(3)

We ensure `set_progress` is run on the *wx* thread, by decorating it with :any:`xloil.gui.wx.wx_thread`.

wxPython Thread-safety
______________________

The :any:`xloil.gui.wx.wx_thread` function behaves the same as `Qt_thread` described
in :ref:`xlOil_Python/CustomGUI:Qt Thread-safety`. 


Task Pane Events
================

Custom task panes have three events which can be handled by defining methods in the subclass of 
:any:`xloil.gui.CustomTaskPane` used to create the pane. The callbacks occur on Excel's main thread.
The events are:

::

    def on_docked(self):
        # Called when the user docks or undocks the pane. The dock position is in 'self.position'
        ...

    def on_visible(self, state: bool):
        # Called when the user closes/shows the pane with the new visibility in 'state'
        ...

    def on_destroy(self):
        # Called just before the pane is destroyed when the parent window is closed
        super().on_destroy() # Important!
        ...

Task Pane registry
==================

The created task panes are automatically stored in a registry so there is no need to hold a
reference to them.  Task panes are attached by default to the active window and it is possible to 
have multiple windows per open workbook.  xlOil will free the panes when the parent workbook ora
addin closes.

We can search the registry by name for a task pane without having the :obj:`xloil.ExcelGUI` object:

::

    pane = xloil.gui.find_task_pane("MyPane")

By default, xlOil looks for a pane attached to the active window, but this can be changed with
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
