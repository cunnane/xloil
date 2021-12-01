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
at the end of the `with` context.


Ribbon
------

xlOil allows dynamic creation of Excel Ribbon components. See :any:`concepts-ribbon` for 
background.  To customise the ribbon, run the following:

::

    ribbon = xlo.ExcelUI(r'''<customUI xmlns=...>....</<customUI>''', 
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

The `MyImageLoader` function will be called with the argument `icon.jpg` as be expected to return a 
*PIL Image*.

Instead of a dictionary, the `mapper` object can be a function which takes any string and returns a 
callback handler.

The ``ribbon`` object returned above is actually a handle to the COM addin created to support
the ribbon customisation.  If the object goes out of scope and is deleted by python or if you call 
``ribbon.disconnect()``, the add-in is unloaded along with the ribbon customisation.

See :doc:`Example` for an example of ribbon customisation.

Custom Task Panes
-----------------

`Custom task panes <https://docs.microsoft.com/en-us/visualstudio/vsto/custom-task-panes>` are user 
interface panels that are typically docked to one side of a window in Excel application.

Custom task panes can be created using the `ExcelUI` object. There is no need to create a ribbon as
well, but task panes are normally opened using a ribon button.

::

    excelui = xlo.ExcelUI(...)
    frame = add_task_pane('MyPane')

To populate the frame, we use a python GUI toolkit (currently only Qt is supported) to draw a pane

::

    from PyQt5.QtWidgets import QWidget
    class MyTaskPane(QWidget):
        ...

    pane = QtThreadTaskPane(frame, MyTaskPane)

To make Qt happy (and not abort), all GUI interaction, other than emitting signals must take place 
in the same thread. `QtThreadTaskPane` creates a dedicated Qt GUI thread for this purpose.