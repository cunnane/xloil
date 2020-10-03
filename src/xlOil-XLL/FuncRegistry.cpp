#include "FuncRegistry.h"
#include <xlOil/Register.h>
#include <xlOil/ExcelCall.h>

#include <xlOil/ExcelObj.h>
#include <xlOil/StaticRegister.h>
#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <xlOil/State.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOilHelpers/PEHelper.h>

#include <codecvt>
#include <map>
#include <filesystem>
namespace fs = std::filesystem;

using std::vector;
using std::shared_ptr;
using std::unique_ptr;
using std::string;
using std::wstring;
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

    const wchar_t* theCoreDllName;

    static int registerWithExcel(
      const shared_ptr<const FuncInfo>& info, 
      const char* entryPoint, 
      const wchar_t* moduleName)
    {
      auto numArgs = info->args.size();
      int opts = info->options;

      // Build the argument type descriptor 
      string argTypes;

      // Set function option prefixes
      if (opts & FuncInfo::ASYNC)
        argTypes += ">X"; // We choose the first argument as the async handle
      else if (opts & FuncInfo::COMMAND)
        argTypes += '>';  // Commands always return void - sensible?
      else               
        argTypes += 'Q';  // Other functions return an XLOPER

      // Arg type Q is XLOPER12 values/arrays
      for (auto& arg : info->args)
        argTypes += arg.allowRange ? 'U' : 'Q';

      // Set function option suffixes
      // TODO: check for invalid combinations
      if (opts & FuncInfo::VOLATILE)
        argTypes += '!';
      else if (opts & FuncInfo::MACRO_TYPE)
        argTypes += '#';
      else if (opts & FuncInfo::THREAD_SAFE)
        argTypes += '$';

      // Concatenate argument names, adding optional indicator if required
      wstring argNames;
      for (auto x : info->args)
        if (x.optional)
          argNames.append(formatStr(L"[%s],", x.name.c_str()));
        else
          argNames.append(x.name).append(L",");

      if (numArgs > 0)
        argNames.pop_back();  // Delete final comma

      const bool truncatedArgNames = argNames.length() > 255;
      if (truncatedArgNames)
      {
        XLO_INFO(L"Excel does not support a concatenated argument name length of "
          "more than 255 chars (including commans). Truncating for function '{0}'", info->name);
        argNames.resize(255);
      }

      // Build argument help strings. If we had to truncate the arg name string
      // add the arg names to the argument help string
      vector<wstring> argHelp;
      if (truncatedArgNames)
        for (auto x : info->args)
          argHelp.emplace_back(fmt::format(L"({0}) {1}", x.name, x.help));
      else
        for (auto x : info->args)
          argHelp.emplace_back(x.help);

      // Pad the last arg help with a couple of spaces to workaround an Excel bug
      if (numArgs > 0 && !argHelp.back().empty())
        argHelp.back() += L"  ";

      // Truncate argument help strings to 255 chars
      for (auto& h : argHelp)
        if (h.size() > 255)
        {
          XLO_INFO(L"Excel does not support argument help strings longer than 255 chars. "
            "Truncating for function '{0}'", info->name);
          h.resize(255);
        }

      // Set the function type
      int macroType = 1;
      if (opts & FuncInfo::COMMAND)
        macroType = 2;
      else if (opts & FuncInfo::HIDDEN)
        macroType = 0;

      // Function help string. Yup, more 255 char limits, those MS folks are terse
      auto truncatedHelp = info->help;
      if (info->help.length() > 255)
      {
        XLO_INFO(L"Excel does not support help strings longer than 255 chars. "
          "Truncating for function '{0}'", info->name);
        truncatedHelp.assign(info->help.c_str(), 255);
        truncatedHelp[252] = '.'; truncatedHelp[253] = '.'; truncatedHelp[254] = '.';
      }

      // TODO: this copies the excelobj
      XLO_DEBUG(L"Registering \"{0}\" at entry point {1} with {2} args", 
        info->name, utf8ToUtf16(entryPoint), numArgs);

      auto registerId = callExcel(xlfRegister,
        moduleName, 
        entryPoint, 
        argTypes, 
        info->name, 
        argNames,
        macroType, 
        info->category, 
        nullptr, nullptr, 
        truncatedHelp.empty() ? info->help : truncatedHelp,
        unpack(argHelp));
      if (registerId.type() != ExcelType::Num)
        XLO_THROW(L"Register '{0}' failed", info->name);
      return registerId.toInt();
    }

    void throwIfPresent(const wstring& name) const
    {
      if (theRegistry.find(name) != theRegistry.end())
        XLO_THROW(L"Function {0} already registered", name);
    }

  public:
    RegisteredFuncPtr add(const shared_ptr<const FuncSpec>& spec)
    {
      auto& name = spec->info()->name;
      throwIfPresent(name);

      return theRegistry.emplace(name, spec->registerFunc()).first->second;
    }

    bool remove(const shared_ptr<RegisteredFunc>& func)
    {
      if (func->deregister())
      {
        theRegistry.erase(func->info()->name);
        return true;
      }
      return false;
    }

    void clear()
    {
      for (auto f : theRegistry)
        const_cast<RegisteredFunc&>(*f.second).deregister();
      theRegistry.clear();
      // theCodePtr = theCodeCave;
    }

    auto find(const wchar_t* name)
    {
      auto found = theRegistry.find(name);
      return found != theRegistry.end() ? found->second : RegisteredFuncPtr();
    }

  private:
    FunctionRegistry()
    {
      theCoreDllName = State::coreName();
    }

    map<wstring, RegisteredFuncPtr> theRegistry;
  };

  RegisteredFunc::RegisteredFunc(const shared_ptr<const FuncSpec>& spec)
    : _spec(spec)
  {}

  RegisteredFunc::~RegisteredFunc()
  {
    deregister();
  }

  bool RegisteredFunc::deregister()
  {
    if (_registerId == 0)
      return false;

    auto& name = info()->name;
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

    // SetExcel12EntryPt is automatically created by xlcall.cpp, but is only used for
    // clusters, which we aren't supporting at this current time.
    auto arbitraryFunction = decorateCFunction("SetExcel12EntryPt", 1);
    auto[tempRegId, retVal] = tryCallExcel(
      xlfRegister, FunctionRegistry::get().theCoreDllName, arbitraryFunction.c_str(), "I", name, nullptr, 2);
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
    return _spec->info();
  }
  const std::shared_ptr<const FuncSpec>& RegisteredFunc::spec() const
  {
    return _spec;
  }
  bool RegisteredFunc::reregister(const std::shared_ptr<const FuncSpec>& /*other*/)
  {
    return false;
  }

  class RegisteredStatic : public RegisteredFunc
  {
  public:
    RegisteredStatic(const std::shared_ptr<const StaticSpec>& spec)
      : RegisteredFunc(spec)
    {
      auto& registry = FunctionRegistry::get();
      _registerId = registry.registerWithExcel(
        spec->info(), 
        decorateCFunction(spec->_entryPoint.c_str(), spec->info()->numArgs()).c_str(), 
        spec->_dllName.c_str());
    }
  };

  std::shared_ptr<RegisteredFunc> StaticSpec::registerFunc() const
  {
    return make_shared<RegisteredStatic>(
      std::static_pointer_cast<const StaticSpec>(this->shared_from_this()));
  }


 
  RegisteredFuncPtr registerFunc(const std::shared_ptr<const FuncSpec>& spec) noexcept
  {
    try
    {
      return FunctionRegistry::get().add(spec);
    }
    catch (std::exception& e)
    {
      XLO_ERROR("Failed to register func {0}: {1}",
        utf16ToUtf8(spec->info()->name.c_str()), e.what());
      return RegisteredFuncPtr();
    }
  }

  int
    registerFuncRaw(
      const std::shared_ptr<const FuncInfo>& info,
      const char* entryPoint,
      const wchar_t* moduleName)
  {
    auto& registry = FunctionRegistry::get();
    return registry.registerWithExcel(info, entryPoint, moduleName);
  }

  RegisteredFuncPtr findRegisteredFunc(const wchar_t * name)
  {
    return FunctionRegistry::get().find(name);
  }
 
  bool deregisterFunc(const shared_ptr<RegisteredFunc>& ptr)
  {
    return FunctionRegistry::get().remove(ptr);
  }
}