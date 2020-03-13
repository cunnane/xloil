#include "FuncRegistry.h"
#include <xlOil/Register.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Events.h>
#include "PEHelper.h"
#include "ExcelObj.h"
#include <xlOil/Log.h>
#include <xlOil/Utils.h>
#include <xlOil/EntryPoint.h>
#include <xlOil/AsyncHelper.h>
#include "Thunker.h"
#include <unordered_set>
#include <codecvt>
#include <future>
#include <map>
#include <filesystem>
namespace fs = std::filesystem;

using std::vector;
using std::shared_ptr;
using std::unique_ptr;
using std::string;
using std::wstring;
using std::unordered_set;
using std::map;
using std::make_shared;
using namespace msxll;

namespace xloil
{
  XLOIL_EXPORT FuncInfo::~FuncInfo()
  {
  }

  XLOIL_EXPORT bool FuncInfo::operator==(const FuncInfo & that) const
  {
    return name == that.name && help == that.help && category == that.category
      && options == that.options && std::equal(args.begin(), args.end(), that.args.begin(), that.args.end());
  }
}

namespace xloil
{
  class FunctionRegistry
  {
  public:
    static FunctionRegistry& get() {
      static FunctionRegistry instance;
      return instance;
    }

    // TODO: We can allocate wipthin our DLL's address space by using
    // NtAllocateVirtualMemory or VirtualAlloc with MEM_TOP_DOWN
    static char theCodeCave[16384 * 2];

    /// <summary>
    /// The next available spot in our code cave
    /// </summary>
    static char* theCodePtr;

    ExcelObj theCoreDllName;
    ExcelObj theXllName;

    template <class TCallback>
    auto callBuildThunk(
      TCallback callback,
      const void* contextData,
      const size_t numArgs)
    {
      // TODO: cache thunks with same number of args and callback?

      const size_t codeBufferSize = sizeof(theCodeCave) + theCodeCave - theCodePtr;
      size_t codeBytesWritten;
      auto* thunk = buildThunk(callback, contextData, numArgs,
        theCodePtr, codeBufferSize, codeBytesWritten);

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
    const char* hookEntryPoint(const FuncInfo&, const void* thunk)
    {
      auto* stubName = theExportTable->getName(theFirstStub);

      // Hook the thunk by modifying the export address table
      XLO_DEBUG("Hooking thunk to name {0}", stubName);
      
      theExportTable->hook(theFirstStub, (void*)thunk);

#ifdef _DEBUG
      // Check the thunk is hooked to Windows' satisfaction
      void* procNew = GetProcAddress((HMODULE)coreModuleHandle(), stubName);
      XLO_ASSERT(procNew == thunk);
#endif

      return stubName;
    }

    int registerWithExcel(shared_ptr<const FuncInfo> info, const char* entryPoint, const ExcelObj& moduleName)
    {
      auto numArgs = info->args.size();
      int opts = info->options;

      string argTypes;

      if (opts & FuncInfo::ASYNC)
        argTypes += ">X"; // We choose the first argument as the async handle
      else if (opts & FuncInfo::COMMAND)
        argTypes += '>';  // Commands always return void - sensible?
      else               
        argTypes += 'Q';  // Other functions return an XLOPER

      // Arg type Q is XLOPER12 values/arrays
      for (auto& arg : info->args)
        argTypes += arg.allowRange ? 'U' : 'Q';

      // TODO: check for invalid combinations
      if (opts & FuncInfo::VOLATILE)
        argTypes += '!';
      else if (opts & FuncInfo::MACRO_TYPE)
        argTypes += '#';
      else if (opts & FuncInfo::THREAD_SAFE)
        argTypes += '$';
 
      vector<wstring> argHelp;
      wstring argNames;
      for (auto x : info->args)
      {
        argNames.append(x.name).append(L",");
        argHelp.emplace_back(x.help);
      }
      
      if (numArgs > 0)
      { 
        // Delete final comma
        argNames.pop_back();

        // Pad the last arg help with a couple of spaces to workaround an Excel bug
        if (!argHelp.back().empty())
          argHelp.back() += L"  ";
      }

      auto macroType = (opts & FuncInfo::COMMAND) ? 2 : 1;

      // TODO: this copies the excelobj
      XLO_DEBUG(L"Registering \"{0}\" at entry point {1} with {2} args", info->name, utf8ToUtf16(entryPoint), numArgs);
      auto registerId = callExcel(xlfRegister,
        moduleName, 
        entryPoint, 
        argTypes, 
        info->name, 
        argNames,
        macroType, 
        info->category, 
        nullptr, nullptr, 
        info->help, 
        unpack(argHelp));

      return registerId.toInt();
    }

    RegisteredFuncPtr addToRegistry(
      const shared_ptr<const FuncInfo>& info, 
      int registerId, 
      shared_ptr<void> context, 
      void* thunk, 
      size_t thunkSize)
    {
      return theRegistry.emplace(info->name, new RegisteredFunc(info, registerId, context, thunk, thunkSize))
        .first->second;
    }

    void throwIfPresent(const wstring& name) const
    {
      if (theRegistry.find(name) != theRegistry.end())
        XLO_THROW(L"Function {0} already registered", name);
    }

  public:
    RegisteredFuncPtr add(const shared_ptr<const FuncInfo>& info, RegisterCallback callback, const std::shared_ptr<void>& data)
    {
      throwIfPresent(info->name);

      auto[thunk, thunkSize] = callBuildThunk(callback, data.get(), info->numArgs());
      auto* entryPoint = hookEntryPoint(*info, thunk);

      auto id = registerWithExcel(info, entryPoint, FunctionRegistry::get().theCoreDllName);

      return addToRegistry(info, id, data, thunk, thunkSize);
    }

    RegisteredFuncPtr add(const shared_ptr<const FuncInfo>& info, AsyncCallback callback, const std::shared_ptr<void>& data)
    {
      throwIfPresent(info->name);

      // Patch up funcinfo in case user didn't set ASYNC flag
      const_cast<FuncInfo&>(*info).options |= FuncInfo::ASYNC;

      auto[thunk, thunkSize] = callBuildThunk(callback, data.get(), info->numArgs());
      auto* entryPoint = hookEntryPoint(*info, thunk);

      auto id = registerWithExcel(info, entryPoint, FunctionRegistry::get().theCoreDllName);

      return addToRegistry(info, id, data, thunk, thunkSize);
    }

    RegisteredFuncPtr add(const shared_ptr<const FuncInfo>& info, const char* entryPoint, const wchar_t* moduleName)
    {
      throwIfPresent(info->name);

      auto id = registerWithExcel(info, entryPoint, ExcelObj(moduleName));

      return addToRegistry(info, id, shared_ptr<void>(), nullptr, 0);
    }

    bool remove(const shared_ptr<RegisteredFunc>& func)
    {
      if (func->deregister())
      {
        theRegistry.erase(func->info()->name);
        // Note this DOES NOT recover the space used for thunks, so we make a note
        _freeThunksAvailable = true;
        return true;
      }
      return false;
    }

    bool compactThunks()
    {
      if (!_freeThunksAvailable)
        return false;
      // TODO: clear and reregister all functions!  Return true if success
      // Or just allocate each thunk with the NtAlloc thingy?
      return false;
    }

    void clear()
    {
      for (auto f : theRegistry)
        const_cast<RegisteredFunc&>(*f.second).deregister();
      theRegistry.clear();
      theCodePtr = theCodeCave;
    }

  private:
    FunctionRegistry()
    {
      theCoreDllName = ExcelObj(theCoreName());
      theXllName = ExcelObj(fs::path(theXllPath()).filename().wstring());
      theExportTable.reset(new DllExportTable((HMODULE)coreModuleHandle()));
      theFirstStub = theExportTable->findOffset(XLO_STR(XLOIL_STUB_NAME));
      _freeThunksAvailable = false;
    }

    map<wstring, RegisteredFuncPtr> theRegistry;
    unique_ptr<DllExportTable> theExportTable;
    size_t theFirstStub;
    
    bool _freeThunksAvailable;
  };

  char FunctionRegistry::theCodeCave[16384 * 2];
  char* FunctionRegistry::theCodePtr = theCodeCave;


  RegisteredFunc::RegisteredFunc(
    const shared_ptr<const FuncInfo>& info,
    int registerId,
    const shared_ptr<void>& context,
    void* thunk,
    size_t thunkSize)
    : _info(info)
    , _registerId(registerId)
    , _context(context)
    , _thunk(thunk)
    , _thunkSize(thunkSize)
  {
  }

  RegisteredFunc::~RegisteredFunc()
  {
    deregister();
  }

  bool RegisteredFunc::deregister()
  {
    if (_registerId == 0)
      return false;

    auto& name = _info->name;
    XLO_DEBUG(L"Deregistering {0}", name);

    auto[result, ret] = tryCallExcel(xlfUnregister, double(_registerId));
    if (ret != msxll::xlretSuccess || result.type() != ExcelType::Bool || !result.toBool())
    {
      XLO_WARN(L"Unregister failed for {0}", name);
      return false;
    }

    // Cunning trick to workaround SetName where function is not removed from wizard
    // by registering a hidden function (i.e. a command) then removing it.  It 
    // doesn't matter which entry point we bind to as long as the function pointer
    // won't be registered as an Excel func.
    // https://stackoverflow.com/questions/15343282/how-to-remove-an-excel-udf-programmatically
    auto[tempRegId, retVal] = tryCallExcel(
      xlfRegister, FunctionRegistry::get().theXllName, "xlAutoOpen", "I", name, nullptr, 2);
    tryCallExcel(xlfSetName, name); // SetName with no arg un-sets the name
    tryCallExcel(xlfUnregister, tempRegId);
    _registerId = 0;
    
    return true;
  }

  int RegisteredFunc::registerId() const
  {
    return _registerId;
  }

  const std::shared_ptr<const FuncInfo>& RegisteredFunc::info() const
  {
    return _info;
  }

  bool RegisteredFunc::reregister(
    const std::shared_ptr<const FuncInfo>& newInfo,
    const std::shared_ptr<void>& newContext)
  {
    XLO_ASSERT(_info->name == newInfo->name);
    if (_thunk && _info->numArgs() == newInfo->numArgs() && _info->options == newInfo->options)
    {
      if (_context != newContext)
      {
        XLO_DEBUG(L"Patching function context for '{0}'", newInfo->name);
        if (!patchThunkData((char*)_thunk, _thunkSize, _context.get(), newContext.get()))
        {
          XLO_ERROR(L"Failed to patch context for '{0}'", newInfo->name);
          return false;
        }
        _context = newContext;
      }

      // If the FuncInfo is identical, no need to re-register
      if (*_info == *newInfo)
      {
        _info = newInfo; // They are already equal by value, but seems the least astonishment approach
        return true;
      }

      // Otherwise re-use the possibly patched thunk

      XLO_DEBUG(L"Reregistering function '{0}'", newInfo->name);
      deregister();
      auto& registry = FunctionRegistry::get();
      auto* entryPoint = registry.hookEntryPoint(*_info, _thunk);
      _registerId = registry.registerWithExcel(_info, entryPoint, registry.theCoreDllName);
      _info = newInfo;
      return true;
    }
    return false;
  }


  namespace
  {
    struct FunctionPrototypeData
    {
      FunctionPrototypeData(const ExcelFuncPrototype& f, shared_ptr<const FuncInfo> i)
        : func(f), info(i)
      {}
      ExcelFuncPrototype func;
      shared_ptr<const FuncInfo> info;
    };

    ExcelObj* launchFunctionObj(void* funcData, const ExcelObj** args)
    {
      auto data = *(FunctionPrototypeData*)funcData;
      return data.func(*data.info, args);
    }

    void launchFunctionObjAsync(void* funcData, const ExcelObj* asyncHandle, const ExcelObj** args)
    {
      try
      {
        auto data = *(FunctionPrototypeData*)funcData;
        auto nArgs = data.info->numArgs();

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
            return data.func(*data.info, &argsPtr[0]);
          }, 
          asyncHandle);

        // Very simple with no cancellation
        std::thread go(functor, 0);
        go.detach();
      }
      catch (...)
      {
      }
    }
  }
  RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, RegisterCallback callback, const std::shared_ptr<void>& data) noexcept
  {
    try
    {
      return FunctionRegistry::get().add(info, callback, data);
    }
    catch (std::exception& e)
    {
      XLO_ERROR("Failed to register func {0}: {1}", 
        utf16ToUtf8(info->name.c_str()), e.what());
      return RegisteredFuncPtr();
    }
  }

  RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, AsyncCallback callback, const std::shared_ptr<void>& data) noexcept
  {
    try
    {
      return FunctionRegistry::get().add(info, callback, data);
    }
    catch (std::exception& e)
    {
      XLO_ERROR("Failed to register func {0}: {1}",
        utf16ToUtf8(info->name.c_str()), e.what());
      return RegisteredFuncPtr();
    }
  }
  RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, const char* functionName, const wchar_t* moduleName) noexcept
  {
    try
    {
      return FunctionRegistry::get().add(info, functionName, moduleName);
    }
    catch (std::exception& e)
    {
      XLO_ERROR("Failed to register func {0} in module {1}", 
        utf16ToUtf8(info->name.c_str()), utf16ToUtf8(moduleName), e.what());
      return RegisteredFuncPtr();
    }
  }

  RegisteredFuncPtr 
    registerFunc(const std::shared_ptr<const FuncInfo>& info, const ExcelFuncPrototype & f) noexcept
  {
    if ((info->options & FuncInfo::ASYNC) != 0)
      return registerFunc(info, &launchFunctionObjAsync, shared_ptr<void>(new FunctionPrototypeData(f, info)));
    else
      return registerFunc(info, &launchFunctionObj, shared_ptr<void>(new FunctionPrototypeData(f, info)));
  }

  bool deregisterFunc(const shared_ptr<RegisteredFunc>& ptr)
  {
    return FunctionRegistry::get().remove(ptr);
  }

  namespace
  {
    struct RegisterMe
    {
      RegisterMe()
      {
        static auto handler = xloil::Event_AutoClose() += []() { FunctionRegistry::get().clear(); };
      }
    } theInstance;
  }
}