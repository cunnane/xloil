======================
xlOil C++ Static XLLs
======================

As explained in :any:`GettingStarted`, it is possible to create an XLL which statically links xlOil.
This allows an all-in-on XLL without worrying about DLL search paths.  We expand on the example 
in :any:`GettingStarted` to illustrate some of the other static XLL features:

.. highlight:: c++

::

    #include <xloil/xlOil.h>
    #include <xloil/XllEntryPoint.h>
    using namespace xloil;

    struct MyAddin
    {
        MyAddin()
        {
            // This constructor is called by Excel's AutoOpen

            xllOpenComCall([this]()
            {
            // If we need to do some COM stuff on startup, like register a Ribbon,
            // it needs to go in this delayed COM callback.
            });
        }

        ~MyAddin()
        {
            // This destructor is called by Excel's AutoClose
        }

        
        static wstring addInManagerInfo()
        {
            // The string returned here is displayed in Excel's addin options window
            // Note the function is static: it can be called before AutoOpen.
            return wstring(L"My Addin Name");
        }
    };
    XLO_DECLARE_ADDIN(MyAddin);

