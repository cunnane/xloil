#include <xlOil/StaticRegister.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOil/StringUtils.h>
#include <xlOil/FuncSpec.h>
#include <xlOil/Throw.h>
#include <xlOil/ExcelCall.h>
#include <filesystem>

using std::vector;
using std::make_shared;
using std::wstring;

namespace xloil
{
  detail::FuncInfoBuilderBase::FuncInfoBuilderBase(size_t nArgs, const int* types)
    : _info(new FuncInfo())
    , _iArg(0)
  {
    _info->args.resize(nArgs);
    for (auto i = 0; i < nArgs; ++i)
      _info->args[i].type = types[i];
  }

  std::shared_ptr<FuncInfo> detail::FuncInfoBuilderBase::getInfo()
  {
    using namespace std::string_literals;

    auto nArgs = _info->args.size();

    for (;_iArg < nArgs; ++_iArg)
      _info->args[_iArg].name = fmt::format(L"Arg_{}", _iArg);

    return _info;
  }

  std::list<FuncRegistrationMemo>& getFuncRegistryQueue()
  {
    static std::list<FuncRegistrationMemo> theQueue;
    return theQueue;
  }

  XLOIL_EXPORT FuncRegistrationMemo& createRegistrationMemo(
    const char* entryPoint_, size_t nArgs, const int* types)
  {
    getFuncRegistryQueue().emplace_back(entryPoint_, nArgs, types);
    return getFuncRegistryQueue().back();
  }

  std::vector<std::shared_ptr<const FuncSpec>>
    processRegistryQueue(const wchar_t* moduleName)
  {
    std::vector<std::shared_ptr<const FuncSpec>> result;
    auto& queue = getFuncRegistryQueue();
    for (auto f : queue)
      result.emplace_back(make_shared<const StaticSpec>(
        f.getInfo(), moduleName, f.entryPoint));
    
    queue.clear();
    return result;
  }

  std::vector<std::shared_ptr<const RegisteredFunc>>
    registerStaticFuncs(const wchar_t* moduleName, std::wstring& errors)
  {
    const auto specs = processRegistryQueue(moduleName);
    std::vector<std::shared_ptr<const RegisteredFunc>> result;
    for (auto& spec : specs)
      try
    {
      result.emplace_back(spec->registerFunc());
    }
    catch (const std::exception& e)
    {
      errors += fmt::format(L"{0}: {1}\n", spec->name(), utf8ToUtf16(e.what()));
    }
    return result;
  }
}