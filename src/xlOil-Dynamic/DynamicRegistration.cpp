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
#include <xlOil-Dynamic/ExternalRegionAllocator.h>
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
  namespace
  {
    constexpr char* XLOIL_STUB_NAME_STR = XLO_STR(XLOIL_STUB_NAME);

    class PageUnlock
    {
    public:
      PageUnlock(void* address, size_t size)
        : _address(address)
        , _size(size)
      {
        if (!VirtualProtect(address, size, PAGE_READWRITE, &_permission))
          XLO_THROW(Helpers::writeWindowsError());
      }
      ~PageUnlock()
      {
        VirtualProtect(_address, _size, _permission, &_permission);
      }
    private:
      void* _address;
      size_t _size;
      DWORD _permission;
    };

    class ThunkHolder
    {
      unique_ptr<DllExportTable> theExportTable;
      int theFirstStub;

      ThunkHolder()
        : theExportTable(new DllExportTable((HMODULE)State::coreModuleHandle()))
        , theAllocator(theExportTable->imageBase(), (BYTE*)theExportTable->imageBase() + DWORD(-1))
      {
        theCoreDllName = State::coreDllName();
        theFirstStub = theExportTable->findOrdinal(
          decorateCFunction(XLOIL_STUB_NAME_STR, 0).c_str());
        if (theFirstStub < 0)
          XLO_THROW("Could not find xlOil stub");
      }

    public:
      const wchar_t* theCoreDllName;
      ExternalRegionAllocator theAllocator;

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
        ThunkWriter writer(callback, contextData, numArgs, hasReturnVal);

        auto codeBytesNeeded = writer.codeSize();

        // We use a custom allocator for the thunks, which must have
        // addresses in the range to be [imageBase, imageBase + DWORD_MAX]
        // described in the DLL export table. Using VirtualAlloc with
        // MEM_TOP_DOWN for some reason is not guaranteed to return
        // addresses above imageBase.
        auto* thunk = theAllocator.alloc((unsigned)codeBytesNeeded);

        DWORD dummy;
        if (!VirtualProtect(thunk, codeBytesNeeded, PAGE_READWRITE, &dummy))
          XLO_THROW(Helpers::writeWindowsError());

        // TODO: compact the alloc if codeBytesWritten < codeBytesNeeded?
        auto codeBytesWritten = writer.writeCode((char*)thunk, codeBytesNeeded);

        // It's good security practice to remove write permissions if we're
        // giving execute permissions, which the thunk code clearly requires
        if (!VirtualProtect(thunk, codeBytesNeeded, PAGE_EXECUTE_READ, &dummy))
          XLO_THROW(Helpers::writeWindowsError());

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
  }

  class RegisteredCallback : public RegisteredWorksheetFunc
  {
  public:
    RegisteredCallback(
      const shared_ptr<const DynamicSpec>& spec, 
      const void* callback,
      const bool callbackHasReturn)
      : RegisteredWorksheetFunc(spec)
    {
      auto& registry = ThunkHolder::get();
      auto[thunk, thunkSize] = registry.callBuildThunk(
        callback, spec->_context.get(), spec->info()->numArgs(), callbackHasReturn);
      _thunk = thunk;
      _thunkSize = thunkSize;
      _registerId = doRegister();
    }

    ~RegisteredCallback()
    {
      ThunkHolder::get().theAllocator.free(_thunk);
    }

    int doRegister() const
    {
      auto& registry = ThunkHolder::get();

      // Point a suitable entry point at our thunk and get its name
      XLO_DEBUG(L"Hooking thunk for {0}", info()->name);
      auto entryPoint = registry.hookEntryPoint(_thunk);
      
      return registerFuncRaw(info(), entryPoint.c_str(), registry.theCoreDllName);
    }

    virtual bool reregister(const std::shared_ptr<const WorksheetFuncSpec>& other) override
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
          PageUnlock unlockPage(_thunk, _thunkSize);
          auto didPatch = patchThunkData((char*)_thunk, _thunkSize, context.get(), newContext.get());
          if (!didPatch)
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

  XLOIL_EXPORT shared_ptr<RegisteredWorksheetFunc> DynamicSpec::registerFunc() const
  {
    return make_shared<RegisteredCallback>(
      static_pointer_cast<const DynamicSpec>(this->shared_from_this()), _callback, _hasReturn);
  }

  namespace
  {
    template<class TRet>
    TRet invokeLambda(
      const LambdaSpec<TRet>* data,
      const ExcelObj** args) noexcept
    {
      try
      {
        return data->function(*data->info(), args);
      }
      catch (const std::exception& e)
      {
        if constexpr (std::is_same_v<TRet, ExcelObj*>)
          return returnValue(e);
        else
        {
          XLO_WARN(e.what());
          if constexpr (std::is_same_v<TRet, int>)
            return 0;
        }
      }
    }
  }

  std::shared_ptr<RegisteredWorksheetFunc> LambdaSpec<ExcelObj*>::registerFunc() const
  {
    auto thisPtr = std::static_pointer_cast<const LambdaSpec<ExcelObj*>>(this->shared_from_this());
    auto thatPtr = make_shared<DynamicSpec>(info(), &invokeLambda<ExcelObj*>, thisPtr);
    return thatPtr->registerFunc();
  }
  std::shared_ptr<RegisteredWorksheetFunc> LambdaSpec<int>::registerFunc() const
  {
    auto thisPtr = std::static_pointer_cast<const LambdaSpec<int>>(this->shared_from_this());
    auto thatPtr = make_shared<DynamicSpec>(info(), &invokeLambda<int>, thisPtr);
    return thatPtr->registerFunc();
  }
  std::shared_ptr<RegisteredWorksheetFunc> LambdaSpec<void>::registerFunc() const
  {
    auto thisPtr = std::static_pointer_cast<const LambdaSpec<void>>(this->shared_from_this());
    auto thatPtr = make_shared<DynamicSpec>(info(), &invokeLambda<void>, thisPtr);
    return thatPtr->registerFunc();
  }
}