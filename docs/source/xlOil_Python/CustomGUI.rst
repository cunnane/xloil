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

The ``mapper`` dictionary links callbacks named in the ribbon XML to python functions. Each
handler should take a single ``RibbonControl`` argument which describes the control which raised 
the callback.  

Alteratively, the `mapper` can be a function which takes any string and returns a callback
handler which takes a ``RibbonControl`` argument.

The ``ribbon`` object returned above is actually a handle to the COM addin created to support
the ribbon customisation.  If the object goes out of scope and is deleted by python or if you call 
``ribbon.disconnect()``, the add-in is unloaded along with the ribbon customisation.

See :doc:`Example` for an example of ribbon customisation.

