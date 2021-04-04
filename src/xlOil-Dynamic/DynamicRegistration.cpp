#include <xloil/DynamicRegister.h>
#include <xloil/ExcelObj.h>
#include <xlOil/Register.h>
#include <xlOil/FuncSpec.h>
#include <xlOil/State.h>
#include <xlOil/Log.h>
#include <xlOil/Throw.h>
#include <xlOil/StaticRegister.h>
#include <xloil/Loaders/EntryPoint.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOil-Dynamic/PEHelper.h>
#include <xlOil-Dynamic/Thunker.h>
#include <xlOil/Preprocessor.h>
#include <xlOil/Async.h>

using std::vector;
using std::shared_ptr;
using std::unique_ptr;
using std::string;
using std::wstring;
using std::make_shared;
using namespace msxll;
using std::static_pointer_cast;

#define XLOIL_STUB_NAME xloil_stub

extern "C"  __declspec(dllexport) void* __stdcall XLOIL_STUB_NAME()
{
  return nullptr;
}

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
      theCoreDllName = State::coreDllName();
      theExportTable.reset(new DllExportTable((HMODULE)State::coreModuleHandle()));
      theFirstStub = theExportTable->findOrdinal(
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

    auto callBuildThunk(
      const void* callback,
      const void* contextData,
      const size_t numArgs,
      const bool hasReturnVal)
    {
      // TODO: cache thunks with same number of args and callback?

      const size_t codeBufferSize = sizeof(theCodeCave) + theCodeCave - theCodePtr;
      size_t codeBytesWritten;
#ifdef _WIN64
      auto* thunk = buildThunkLite(callback, contextData, numArgs, hasReturnVal,
        theCodePtr, codeBufferSize, codeBytesWritten);
#else
      auto* thunk = buildThunk(callback, contextData, numArgs, hasReturnVal,
        theCodePtr, codeBufferSize, codeBytesWritten);
#endif
      XLO_ASSERT(thunk == (void*)theCodePtr);
      theCodePtr += codeBytesWritten;
      return std::make_pair(thunk, codeBytesWritten);
    }

    /// <summary>
    /// Locates a suitable entry point in our DLL and hooks the specifed thunk to it
    /// </summary>
    /// <returns>The name of the entry point selected</returns>
    auto hookEntryPoint(const void* thunk)
    {
      // Hook the thunk by modifying the export address table
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


  class RegisteredCallback : public RegisteredWorksheetFunc
  {
  public:
    RegisteredCallback(
      const shared_ptr<const DynamicSpec>& spec)
      : RegisteredWorksheetFunc(spec)
    {
      auto& registry = ThunkHolder::get();
      auto[thunk, thunkSize] = registry.callBuildThunk(
        spec->_callback, spec->_context.get(), spec->info()->numArgs(), spec->_hasReturn);
      _thunk = thunk;
      _thunkSize = thunkSize;
      _registerId = doRegister();
    }

    int doRegister() const
    {
      auto& registry = ThunkHolder::get();

      // Point a suitable entry point at our thunk and get its name
      XLO_DEBUG(L"Hooking thunk for {0}", info()->name);
      auto entryPoint = registry.hookEntryPoint(_thunk);
      
      return registerFuncRaw(info(), entryPoint.c_str(), registry.theCoreDllName);
    }

    virtual bool reregister(const std::shared_ptr<const WorksheetFuncSpec>& other)
    {
      auto* thisType = dynamic_cast<const DynamicSpec*>(other.get());
      if (!thisType)
        return false;

      auto& newInfo = other->info();
      auto newContext = thisType->_context;
      auto& context = spec()._context;

      XLO_ASSERT(info()->name == newInfo->name);
      if (_thunk 
        && info()->numArgs() == newInfo->numArgs() 
        && info()->options == newInfo->options)
      {
        const bool infoMatches = *info() == *newInfo;
        const bool contextMatches = context == newContext;

        if (!contextMatches)
        {
          XLO_DEBUG(L"Patching function context for '{0}'", newInfo->name);
          if (!patchThunkData((char*)_thunk, _thunkSize, context.get(), newContext.get()))
          {
            XLO_ERROR(L"Failed to patch context for '{0}'", newInfo->name);
            return false;
          }
        }
        
        // Rewrite spec so new context and info pointers are kept alive
        _spec = other;
    
        // If the FuncInfo is identical, no need to re-register
        if (infoMatches)
          return true;

        // Otherwise do full re-registration
        XLO_DEBUG(L"Reregistering function '{0}'", newInfo->name);
        deregister();
        _registerId = doRegister();
        
        return true;
      }
      return false;
    }

    const DynamicSpec& spec() const
    {
      return static_cast<const DynamicSpec&>(*_spec);
    }

  private:
    void* _thunk;
    size_t _thunkSize;
  };

  shared_ptr<RegisteredWorksheetFunc> DynamicSpec::registerFunc() const
  {
    return make_shared<RegisteredCallback>(
      static_pointer_cast<const DynamicSpec>(
        this->shared_from_this()));
  }

  namespace
  {
    ExcelObj* invokeLambda(
      const LambdaSpec<ExcelObj*>* data,
      const ExcelObj** args) noexcept
    {
      try
      {
        return data->function(*data->info(), args);
      }
      catch (const std::exception& e)
      {
        return returnValue(e);
      }
    }

    void invokeVoidLambda(
      const LambdaSpec<void>* data,
      const ExcelObj** args) noexcept
    {
      try
      {
        data->function(*data->info(), args);
      }
      catch (...)
      {
      }
    }
  }

  std::shared_ptr<RegisteredWorksheetFunc> LambdaSpec<ExcelObj*>::registerFunc() const
  {
    auto thisPtr = std::static_pointer_cast<const LambdaSpec<ExcelObj*>>(this->shared_from_this());
    auto thatPtr = make_shared<DynamicSpec>(info(), &invokeLambda, thisPtr);
    return thatPtr->registerFunc();
  }
  std::shared_ptr<RegisteredWorksheetFunc> LambdaSpec<void>::registerFunc() const
  {
    auto thisPtr = std::static_pointer_cast<const LambdaSpec<void>>(this->shared_from_this());
    auto thatPtr = make_shared<DynamicSpec>(info(), &invokeVoidLambda, thisPtr);
    return thatPtr->registerFunc();
  }
}