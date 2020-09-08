==============================
xlOil Python GUI Customisation
==============================

.. contents::
    :local:

Ribbon
------

xlOil allows dynamic creation of Excel Ribbon components. The ribbon is defined by XML
(surrounded with <customUI> tags) which should be created with a specialised editor, see the 
*Resources* below. Controls in the ribbon interact with user code via callback handlers. 
To customise the ribbon, simply run the following:

::

    ribbon = xlo.create_ribbon(r'''<customUI xmlns=...>....</<customUI>''', 
        handlers={
            'onClick1': run_click,
            'onClick2': run_click,
            'onUpdate': update_text,
        })

The ``handlers`` dictionary links callbacks named in the ribbon XML to python functions. Each
handler should take a single ``RibbonControl`` argument which describes the control which raised 
the callback.

To pass ribbon XML into Excel, xlOil creates a COM-based add-in in addition to the XLL-based 
add-in which loads the xlOil core - you can see this appearing in Excel's add-in list in the 
Excel's Options windows.  The ``ribbon`` object returned above is actually a handle to this
add-in.  If the object goes out of scope and is deleted by python or if you call 
``ribbon.disconnect()``, the add-in is unloaded along with the ribbon customisation.

See :doc:`Example` for an example of ribbon customisation.

Resources:

   * `Microsoft: Overview of the Office Fluent Ribbon <https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/overview-of-the-office-fluent-ribbon>`_
   * `Microsoft: Customizing the Office Fluent Ribbon for Developers <https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338202(v=office.12)>`_
   * `Microsoft: Custom UI XML Markup Specification <https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43>`_
   * `Office RibbonX Editor <https://github.com/fernandreu/office-ribbonx-editor>`_
   * `Ron de Bruin: Ribbon Examples files and Tips <https://www.rondebruin.nl/win/s2/win003.htm>`_

