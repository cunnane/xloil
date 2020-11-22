==============================
xlOil Python GUI Customisation
==============================

.. contents::
    :local:

Ribbon
------

xlOil allows dynamic creation of Excel Ribbon components. See :any:`concepts-ribbon` for 
background.  To customise the ribbon, simply run the following:

::

    ribbon = xlo.create_ribbon(r'''<customUI xmlns=...>....</<customUI>''', 
        mapper={
            'onClick1': run_click,
            'onClick2': run_click,
            'onUpdate': update_text,
        })

The ``mapper`` dictionary (or function) links callbacks named in the ribbon XML to python functions. 
Each handler should have a signature like the following:

::

    def ribbon_callback(ctrl: RibbonControl)
        ...
    def ribbon_callback(ctrl: RibbonControl, arg1, arg2)
        ...
    def ribbon_callback(ctrl: RibbonControl, *args)
        ...    


The ``RibbonControl`` describes the control which raised the callback. The number of additional
arguments is callback dependent.  In addition, the callback may be expected to return a value. 
See the *Resources* in :any:`concepts-ribbon` for a description of the appropriate callback signature.

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

