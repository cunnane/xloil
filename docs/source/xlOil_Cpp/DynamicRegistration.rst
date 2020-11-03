==============================
xlOil C++ Dynamic Registration
==============================

To support registration of functions in other languages, xlOil has the ability to register worksheet
functions generated at runtime.  In C++, you usually can create a static entry point for each function
you want to register using xlOil macros which end up calling `Excel12(xlfRegister, ...)`.  To

::

    auto regId = RegisterLambda(
        [](const FuncInfo& info, const ExcelObj& arg1, const ExcelObj& arg2)
        {
            ...
            return returnValue(...);
        })
        .name("Foobar")
        .help("Does nothing")
        .arg("Arg1")
        .registerFunc();

The Lambda's first argument must be `const FuncInfo&` and then as many `const ExcelObj&` args as
required; it must return a `ExcelObj*`. It must not throw.