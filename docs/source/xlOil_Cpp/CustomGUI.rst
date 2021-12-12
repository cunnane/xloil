==============================
xlOil C++ GUI Customisation
==============================

.. contents::
    :local:

xlOil allows dynamic creation of Excel Ribbon components and custom task panes. See 
:any:`concepts-ribbon` for background. The code differs slightly depending on whether 
you are writing a static XLL or an xlOil plugin.

Creating a ribbon in a static XLL
---------------------------------

.. highlight:: c++

::
    
    void ribbonHandler(const RibbonControl& ctrl) {}
    

    struct MyAddin
    {
        std::shared_ptr<IComAddin> theComAddin;
        MyAddin()
        {
            // We need this call so that xlOil defers execution of the ribbon creation
            // until Excel's COM interface is ready
            xllOpenComCall([this]()
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
    };
    XLO_DECLARE_ADDIN(MyAddin);


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

    
Creating a Custom Task Pane
---------------------------

After a COM addin has been created using the ribbon examples above, a custom task pane
can be added in the following way.

.. highlight:: c++

:: 
    
    std::shared_ptr<ICustomTaskPane> taskPane(theComAddin->createTaskPane(L"xloil"));
    taskPane->setVisible(true);

An option progid can be passed to `createTaskPane` to create a specific COM object 
as the root of the task pane, if omitted, xlOil uses a simple default object.