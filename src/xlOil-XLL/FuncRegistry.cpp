#include "FuncRegistry.h"
#include "Intellisense.h"

#include <xlOil/Register.h>
#include <xlOil/ExcelCall.h>

#include <xlOil/ExcelObj.h>
#include <xlOil/StaticRegister.h>
#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <xlOil/State.h>
#include <xlOil-Dynamic/PEHelper.h>

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

  XLOIL_EXPORT bool FuncInfo::operator==(const FuncInfo& that) const
  {
    return name == that.name && help == that.help && category == that.category
      && options == that.options && std::equal(args.begin(), args.end(), that.args.begin(), that.args.end());
  }
  XLOIL_EXPORT bool FuncInfo::isValid() const
  {
    bool async = false;
    for (auto& arg : args)
      if (arg.type & FuncArg::AsyncHandle)
        async = true;
    // Confirm we don't set more than one of these
    return (((options & FuncInfo::MACRO_TYPE) > 0)
      + ((options & FuncInfo::THREAD_SAFE) > 0)
      + ((options & FuncInfo::COMMAND) > 0)
      + (async) <= 1);
  }
}

namespace xloil
{
  class FunctionRegistry
  {
  public:
    const wchar_t* theCoreDllName;

  private:
    map<wstring, RegisteredFuncPtr> theRegistry;

    FunctionRegistry()
    {
      theCoreDllName = Environment::coreDllName();
    }

    ~FunctionRegistry()
    {
      teardown();
    }

  public:
    static FunctionRegistry& get()
    {
      static FunctionRegistry instance;
      return instance;
    }

    void teardown()
    {
      // If we reach static destruction and still have registered functions is means
      // something was not properly cleaned up during XLL autoClose. We are not in 
      // an XLL context so trying to deregister will fail or even crash Excel.
      for (auto& entry : theRegistry)
        entry.second->forget();
      theRegistry.clear();
    }

    static int registerWithExcel(
      const shared_ptr<const FuncInfo>& info,
      const char* entryPoint,
      const wchar_t* moduleName)
    {
      auto numArgs = info->args.size();
      int opts = info->options;

      if (numArgs > XL_MAX_UDF_ARGS)
        XLO_THROW("Number of positional arguments ({}) exceeds maximum allowed by Excel ({})",
          numArgs, XL_MAX_UDF_ARGS);

      if (!info->isValid())
        XLO_THROW("Invalid combination of function options");

      // Build the argument type descriptor
      string argTypes;
      if (opts & FuncInfo::COMMAND)
        argTypes += 'A';  // Commands always return int
      else if (opts & FuncInfo::ARRAY)
        argTypes += "K%";  // FP12 struct
      else
        argTypes += 'U';  // Otherwise return an XLOPER12 unless overridden below

      int iArg = 0;
      for (auto& arg : info->args)
      {
        if (arg.type & FuncArg::Range)
          argTypes += 'U';  // XLOPER12 values/arrays and Ranges
        else if (arg.type & FuncArg::Obj)
          argTypes += 'Q';  // XLOPER12 values/arrays
        else if (arg.type & FuncArg::Array)
          argTypes += "K%"; // FP12 struct
        else if (arg.type & FuncArg::AsyncHandle)
        {
          argTypes += "X";   // Async return handle
          argTypes[0] = '>'; // Async returns void
        }
        else
          XLO_THROW(L"Internal: Unknown argtype '{}' for arg '{}'", arg.type, arg.name);

        if (arg.type & FuncArg::ReturnVal)
          if (argTypes[0] != 'U')
            XLO_THROW(L"Only one argument can be specified as a return value for arg '{}'", arg.name);
          else if (iArg > 8)
            XLO_THROW(L"Return in-place arg must be in the first 9 for arg '{}'", arg.name);
          else
            argTypes[0] = ('1' + (char)iArg); // Return numbered arg in place
        ++iArg;
      }

      // Set function option suffixes
      if (opts & FuncInfo::VOLATILE)
        argTypes += '!';
      else if (opts & FuncInfo::MACRO_TYPE)
        argTypes += '#';
      else if (opts & FuncInfo::THREAD_SAFE)
      {
        argTypes += '$';
        if (opts & FuncInfo::VOLATILE)
          XLO_THROW("Cannot declare function thread-safe and volatile");
      }

      // Concatenate argument names, adding optional indicator if requested
      wstring argNames;
      for (auto x : info->args)
        if (x.type & FuncArg::Optional)
          argNames.append(formatStr(L"[%s],", x.name.c_str()));
        else
          argNames.append(x.name).append(L",");

      if (numArgs > 0)
        argNames.pop_back();  // Delete final comma

      const bool truncatedArgNames = argNames.length() > 255;
      if (truncatedArgNames)
      {
        XLO_INFO(L"Excel does not support a concatenated argument name length of "
          "more than 255 chars (including commas). Truncating for function '{0}'", info->name);
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
        if (h.size() > XL_ARG_HELP_STRING_MAX_LENGTH)
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
      if (info->help.length() > XL_ARG_HELP_STRING_MAX_LENGTH)
      {
        constexpr auto maxLen = XL_ARG_HELP_STRING_MAX_LENGTH;
        XLO_INFO(L"Excel does not support help strings longer than {1} chars. "
          "Truncating for function '{0}'", info->name, maxLen);
        truncatedHelp.assign(info->help.c_str(), maxLen);
        truncatedHelp[maxLen - 3] = '.'; truncatedHelp[maxLen - 2] = '.'; truncatedHelp[maxLen - 1] = '.';
      }

      // TODO: entrypoint will always be ascii
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

      // We must be in XLL context to register a function, so can call this:
      publishIntellisenseInfo(info);

      return registerId.get<int>();
    }

    void throwIfPresent(const wstring& name) const
    {
      if (theRegistry.find(name) != theRegistry.end())
        XLO_THROW(L"Function {0} already registered", name);
    }

  public:
    RegisteredFuncPtr add(const shared_ptr<const WorksheetFuncSpec>& spec)
    {
      auto& name = spec->info()->name;
      throwIfPresent(name);

      return theRegistry.emplace(name, spec->registerFunc()).first->second;
    }

    void remove(const wchar_t* funcName)
    {
      theRegistry.erase(funcName);
    }

    void clear()
    {
      for (auto f : theRegistry)
        const_cast<RegisteredWorksheetFunc&>(*f.second).deregister();
      theRegistry.clear();
    }

    auto& all()
    {
      return theRegistry;
    }
  };

  RegisteredWorksheetFunc::RegisteredWorksheetFunc(const shared_ptr<const WorksheetFuncSpec>& spec)
    : _spec(spec)
  {}

  RegisteredWorksheetFunc::~RegisteredWorksheetFunc()
  {
    deregister();
  }

  bool RegisteredWorksheetFunc::deregister()
  {
    if (_registerId == 0)
      return false;

    auto& name = info()->name;
    XLO_DEBUG(L"Deregistering {0}", name);

    auto [result, ret] = tryCallExcel(xlfUnregister, double(_registerId));
    if (ret != msxll::xlretSuccess || result.type() != ExcelType::Bool || !result.get<bool>())
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
    auto [tempRegId, retVal] = tryCallExcel(
      xlfRegister, FunctionRegistry::get().theCoreDllName, arbitraryFunction.c_str(), "I", name, nullptr, 2);
    tryCallExcel(xlfSetName, name); // SetName with no arg un-sets the name
    tryCallExcel(xlfUnregister, tempRegId);
    _registerId = 0;

    FunctionRegistry::get().remove(name.c_str());

    return true;
  }

  int RegisteredWorksheetFunc::registerId() const
  {
    return _registerId;
  }

  const std::shared_ptr<const FuncInfo>& RegisteredWorksheetFunc::info() const
  {
    return _spec->info();
  }
  const std::shared_ptr<const WorksheetFuncSpec>& RegisteredWorksheetFunc::spec() const
  {
    return _spec;
  }
  bool RegisteredWorksheetFunc::reregister(const std::shared_ptr<const WorksheetFuncSpec>& /*other*/)
  {
    return false;
  }

  class RegisteredStatic : public RegisteredWorksheetFunc
  {
  public:
    RegisteredStatic(const std::shared_ptr<const StaticWorksheetFunction>& spec)
      : RegisteredWorksheetFunc(spec)
    {
      auto& registry = FunctionRegistry::get();
      _registerId = registry.registerWithExcel(
        spec->info(),
        decorateCFunction(spec->_entryPoint.c_str(), spec->info()->numArgs()).c_str(),
        spec->_dllName.c_str());
    }
  };

  std::shared_ptr<RegisteredWorksheetFunc> StaticWorksheetFunction::registerFunc() const
  {
    try
    {
      return make_shared<RegisteredStatic>(
        std::static_pointer_cast<const StaticWorksheetFunction>(this->shared_from_this()));
    }
    catch (const std::exception& e)
    {
      XLO_THROW("{0}. Error registering '{1}'", e.what(), utf16ToUtf8(this->name()));
    }
  }

  RegisteredFuncPtr registerFunc(const std::shared_ptr<const WorksheetFuncSpec>& spec) noexcept
  {
    try
    {
      return FunctionRegistry::get().add(spec);
    }
    catch (std::exception& e)
    {
      XLO_ERROR(L"Failed to register func {0}: {1}",
        spec->info()->name.c_str(), utf8ToUtf16(e.what()));
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

  const map<wstring, RegisteredFuncPtr>& registeredFuncsByName()
  {
    return FunctionRegistry::get().all();
  }

  void teardownFunctionRegistry()
  {
    FunctionRegistry::get().teardown();
  }
}