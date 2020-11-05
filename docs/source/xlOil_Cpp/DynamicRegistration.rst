==============================
xlOil C++ Dynamic Registration
==============================

To support registration of functions in other languages, xlOil has the ability to register worksheet
functions generated at runtime.  In C++, you usually can create a static entry point for each function
you want to register using xlOil macros which end up calling `Excel12(xlfRegister, ...)`.  To

::

    auto regId = RegisterLambda<>(
        [](const ExcelObj& arg1, const ExcelObj& arg2)
        {
            ...
            return returnValue(...);
        })
        .name("Foobar")
        .help("Does nothing")
        .arg("Arg1")
        .registerFunc();

The lambda's can take as many `const ExcelObj&` args as required. It must return `ExcelObj*`,
but it may throw: xlOil will return the error string.  By specifying `const FuncInfo&` as the 
type of the first argument, the callable will be given a reference to the function registration info
in addition the arguments passed by Excel.

You can dynamically register any function you could register statically, so :any:`SpecialArgs` 
arguments are valid, as well as a void return, for example:

::

    auto regId = RegisterLambda<void>(
        [](const FuncInfo&, const ExcelObj& arg1, const AsyncHandle& handle)
        {
            handle.returnValue(...);
        })
        .name("AsyncFoobar").registerFunc();

The returned register Id is a `shared_ptr` whose destructor unregisters the function, so this must be
kept in scope.

