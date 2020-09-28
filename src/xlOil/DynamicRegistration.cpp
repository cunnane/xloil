#include <xloil/ExcelObj.h>
#include <xlOil/Register.h>
#include <xlOil/FuncSpec.h>
#include <xlOil/State.h>
#include <xlOil/Log.h>
#include <xlOil/Throw.h>
#include <xlOil/StaticRegister.h>
#include <xloil/Loaders/EntryPoint.h>
#include <xlOil/Register/FuncRegistry.h>
#include <xlOilHelpers/PEHelper.h>
#include <xlOilHelpers/Thunker.h>
#include <xlOil/Preprocessor.h>
#include <xlOil/Async.h>

using std::vector;
using std::shared_ptr;
using std::unique_ptr;
using std::string;
using std::wstring;
using std::make_shared;
using namespace msxll;


namespace xloil
{
  constexpr char* XLOIL_STUB_NAME_STR = XLO_STR(XLOIL_STUB_NAME);

  class ThunkHolder
  {
    // TODO: We can allocate within our DLL's address space by using
    // NtAllocateVirtualMemory or VirtualAlloc with MEM_TOP_DOWN
    // Currently this gives space for about 1500 thunks
    static constexpr auto theCaveSize = 16384 * 8u;
    static char theCodeCave[theCaveSize];
    unique_ptr<DllExportTable> theExportTable;
    int theFirstStub;
    
    ThunkHolder()
    {
      theCoreDllName = State::coreName();
      theExportTable.reset(new DllExportTable((HMODULE)State::coreModuleHandle()));
      theFirstStub = theExportTable->findOffset(
        decorateCFunction(XLOIL_STUB_NAME_STR, 0).c_str());
      if (theFirstStub < 0)
        XLO_THROW("Could not find xlOil stub");
    }

    /// <summary>
    /// The next available spot in our code cave
    /// </summary>
    static char* theCodePtr;

  public:
    const wchar_t* theCoreDllName;

    static ThunkHolder& get() {
      static ThunkHolder instance;
      return instance;
    }

    template <class TCallback>
    auto callBuildThunk(
      TCallback callback,
      const void* contextData,
      const size_t numArgs)
    {
      // TODO: cache thunks with same number of args and callback?

      const size_t codeBufferSize = sizeof(theCodeCave) + theCodeCave - theCodePtr;
      size_t codeBytesWritten;
#ifdef _WIN64
      auto* thunk = buildThunkLite(callback, contextData, numArgs,
        theCodePtr, codeBufferSize, codeBytesWritten);
#else
      auto* thunk = buildThunk(callback, contextData, numArgs,
        theCodePtr, codeBufferSize, codeBytesWritten);
#endif
      XLO_ASSERT(thunk == (void*)theCodePtr);
      theCodePtr += codeBytesWritten;
      return std::make_pair(thunk, codeBytesWritten);
    }

    /// <summary>
    /// Locates a suitable entry point in our DLL and hooks the specifed thunk to it
    /// </summary>
    /// <param name="info"></param>
    /// <param name="thunk"></param>
    /// <returns>The name of the entry point selected</returns>
    auto hookEntryPoint(const FuncInfo& info, const void* thunk)
    {
      // Hook the thunk by modifying the export address table
      XLO_DEBUG(L"Hooking thunk for {0}", info.name);

      theExportTable->hook(theFirstStub, (void*)thunk);

      const auto entryPoint = decorateCFunction(XLOIL_STUB_NAME_STR, 0);

#ifdef _DEBUG
      // Check the thunk is hooked to Windows' satisfaction
      void* procNew = GetProcAddress((HMODULE)State::coreModuleHandle(),
        entryPoint.c_str());
      XLO_ASSERT(procNew == thunk);
#endif

      return entryPoint;
    }
  };
  char ThunkHolder::theCodeCave[theCaveSize];
  char* ThunkHolder::theCodePtr = theCodeCave;


  template <class TCallback>
  class RegisteredCallback : public RegisteredFunc
  {
  public:
    RegisteredCallback(
      const std::shared_ptr<const GenericCallbackSpec<TCallback>>& spec)
      : RegisteredFunc(spec)
    {
      auto& registry = ThunkHolder::get();
      auto[thunk, thunkSize] = registry.callBuildThunk(
        spec->_callback, spec->_context.get(), spec->info()->numArgs());
      _thunk = thunk;
      _thunkSize = thunkSize;
      _registerId = doRegister();
    }

    int doRegister() const
    {
      auto& registry = ThunkHolder::get();
      auto entryPoint = registry.hookEntryPoint(*info(), _thunk);
      return registerFuncRaw(info(), entryPoint.c_str(), registry.theCoreDllName);
    }

    virtual bool reregister(const std::shared_ptr<const FuncSpec>& other)
    {
      auto* thisType = dynamic_cast<const GenericCallbackSpec<TCallback>*>(other.get());
      if (!thisType)
        return false;

      auto& newInfo = other->info();
      auto newContext = thisType->_context;
      auto& context = spec()._context;

      XLO_ASSERT(info()->name == newInfo->name);
      if (_thunk && info()->numArgs() == newInfo->numArgs() && info()->options == newInfo->options)
      {
        bool infoMatches = *info() == *newInfo;
        bool contextMatches = context != newContext;

        if (!contextMatches)
        {
          XLO_DEBUG(L"Patching function context for '{0}'", newInfo->name);
          if (!patchThunkData((char*)_thunk, _thunkSize, context.get(), newContext.get()))
          {
            XLO_ERROR(L"Failed to patch context for '{0}'", newInfo->name);
            return false;
          }
        }
        // If the FuncInfo is identical, no need to re-register, note this
        // discards the new funcinfo.
        if (infoMatches)
          return true;

        // Rewrite spec
        _spec = make_shared<GenericCallbackSpec<TCallback>>(newInfo, spec()._callback, newContext);

        // Otherwise re-use the possibly patched thunk
        XLO_DEBUG(L"Reregistering function '{0}'", newInfo->name);
        deregister();
        _registerId = doRegister();
        _spec = other;
        return true;
      }
      return false;
    }

    const GenericCallbackSpec<TCallback>& spec() const
    {
      return static_cast<const GenericCallbackSpec<TCallback>&>(*_spec);
    }

  private:
    void* _thunk;
    size_t _thunkSize;
  };

  std::shared_ptr<RegisteredFunc> GenericCallbackSpec<RegisterCallback>::registerFunc() const
  {
    return make_shared<RegisteredCallback<RegisterCallback>>(
      std::static_pointer_cast<const GenericCallbackSpec<RegisterCallback>>(this->shared_from_this()));
  }

  std::shared_ptr<RegisteredFunc> GenericCallbackSpec<AsyncCallback>::registerFunc() const
  {
    return make_shared<RegisteredCallback<AsyncCallback>>(
      std::static_pointer_cast<const GenericCallbackSpec<AsyncCallback>>(this->shared_from_this()));
  }

  namespace
  {
    ExcelObj* launchFunctionObj(
      LambdaFuncSpec* data,
      const ExcelObj** args) noexcept
    {
      try
      {
        return data->_function(*data->info(), args);
      }
      catch (const std::exception& e)
      {
        return returnValue(e);
      }
    }

  // TODO: this is not used and maybe not that useful!
  class AsyncHolder
  {
  public:
    // No need to copy the data as FuncRegistry will keep this alive
    // Async handle is destroyed by Excel return, so must copy that
    AsyncHolder(std::function<ExcelObj*()> func, const ExcelObj* asyncHandle)
      : _call(func)
      , _asyncHandle(*asyncHandle)
    {
    }
    void operator()(int /*threadId*/) const
    {
      auto* result = _call();
      asyncReturn(_asyncHandle, ExcelObj(*result));
      if (result->xltype & msxll::xlbitDLLFree)
        delete result;
    }
  private:
    std::function<ExcelObj*()> _call;
    ExcelObj _asyncHandle;
  };

  void launchFunctionObjAsync(
    LambdaFuncSpec* data,
    const ExcelObj* asyncHandle,
    const ExcelObj** args) noexcept
  {
    try
    {
      auto nArgs = data->info()->numArgs();

      // Make a shared_ptr so the lambda below can capture it without a copy
      auto argsCopy = make_shared<vector<ExcelObj>>();
      argsCopy->reserve(nArgs);
      std::transform(args, args + nArgs, std::back_inserter(*argsCopy), [](auto* p) {return ExcelObj(*p); });

      auto functor = AsyncHolder(
        [argsCopy, data]()
      {
        std::vector<const ExcelObj*> argsPtr;
        argsPtr.reserve(argsCopy->size());
        std::transform(argsCopy->begin(), argsCopy->end(), std::back_inserter(argsPtr), [](ExcelObj& x) { return &x; });
        return data->_function(*data->info(), &argsPtr[0]);
      },
        asyncHandle);

      //Very simple with no cancellation
      std::thread go(functor, 0);
      go.detach();
    }
    catch (...)
    {
    }
  }
}

std::shared_ptr<RegisteredFunc> LambdaFuncSpec::registerFunc() const
{
  auto copyThis = make_shared<LambdaFuncSpec>(*this);
  if ((info()->options & FuncInfo::ASYNC) != 0)
    return AsyncCallbackSpec(info(), &launchFunctionObjAsync, copyThis).registerFunc();
  else
    return CallbackSpec(info(), &launchFunctionObj, copyThis).registerFunc();
}

}