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
    for (size_t i = 0; i < nArgs; ++i)
      _info->args[i].type = types[i];
  }

  std::shared_ptr<FuncInfo> detail::FuncInfoBuilderBase::getInfo()
  {
    auto nArgs = _info->args.size();

    for (size_t i = 0; i < nArgs; ++i)
      if (_info->args[i].name.empty())
        _info->args[i].name = fmt::format(L"Arg_{}", i);

    return _info;
  }

  namespace detail
  {
    std::list<StaticRegistrationBuilder>& getFuncRegistryQueue()
    {
      static std::list<StaticRegistrationBuilder> theQueue;
      return theQueue;
    }

    StaticRegistrationBuilder& createRegistrationMemo(
      const char* entryPoint_, int funcOpts, size_t nArgs, const int* types)
    {
      getFuncRegistryQueue().emplace_back(entryPoint_, funcOpts, nArgs, types);
      return getFuncRegistryQueue().back();
    }

    std::vector<std::shared_ptr<const WorksheetFuncSpec>>
      processRegistryQueue(const wchar_t* moduleName)
    {
      std::vector<std::shared_ptr<const WorksheetFuncSpec>> result;
      auto& queue = getFuncRegistryQueue();
      for (auto f : queue)
        result.emplace_back(f.writeFuncSpec(moduleName));

      queue.clear();
      return result;
    }

    std::vector<std::shared_ptr<const RegisteredWorksheetFunc>>
      registerStaticFuncs(const wchar_t* moduleName, std::wstring& errors)
    {
      const auto specs = processRegistryQueue(moduleName);
      std::vector<std::shared_ptr<const RegisteredWorksheetFunc>> result;
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
}