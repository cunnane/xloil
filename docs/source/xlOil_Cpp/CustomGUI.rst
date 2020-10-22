==============================
xlOil C++ GUI Customisation
==============================

xlOil allows dynamic creation of Excel Ribbon components. See :any:`concepts-ribbon:` for 
background. The code differs slightly depending on whether you are writing a static XLL or
an xlOil plugin.

Creating a ribbon in a static XLL
---------------------------------

.. highlight:: c++

::
    
    void ribbonHandler(const RibbonControl& ctrl) {}
    std::shared_ptr<IComAddin> theComAddin;

    void xllOpen(void* hInstance)
    {
        // We need this call so that xlOil defers execution of the ribbon creation
        // until Excel's COM interface is ready
        xllOpenComCall([]()
        {
            auto mapper = [=](const wchar_t* name) mutable { return ribbonHandler; };
            theComAddin = makeAddinWithRibbon(
                L"TestXlOil",
                LR"(
                    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
                    ...
                    </customUI>
                )", 
                mapper);
        });
    }


Creating a ribbon in an xlOil plugin
------------------------------------

.. highlight:: c++

:: 

    #incude <xloil/xloil.h>

    void ribbonHandler(const RibbonControl& ctrl) {}
    std::shared_ptr<IComAddin> theComAddin;

    XLO_PLUGIN_INIT(xloil::AddinContext* ctx, const xloil::PluginContext& plugin)
    {
        xloil::linkLogger(ctx, plugin);

        auto mapper = [=](const wchar_t* name) mutable { return ribbonHandler; };
        theComAddin = makeAddinWithRibbon(
            L"TestXlOil",
            LR"(
                <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
                ...
                </customUI>
            )", 
            mapper);

        return 0;
    }