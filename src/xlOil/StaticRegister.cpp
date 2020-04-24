#include "StaticRegister.h"
#include <xlOil/Register/FuncRegistry.h>
#include <xlOilHelpers/StringUtils.h>
#include "FuncSpec.h"
#include "Throw.h"

using std::vector;
using std::make_shared;

namespace xloil
{
  FuncRegistrationMemo::FuncRegistrationMemo(const char* entryPoint_, size_t nArgs)
    : _nArgs(nArgs)
    , entryPoint(entryPoint_)
    , _info(new FuncInfo())
    , _allowRangeAll(false)
  {
    _info->name = utf8ToUtf16(entryPoint_);
  }

  std::shared_ptr<const FuncInfo> FuncRegistrationMemo::getInfo()
  {
    using namespace std::string_literals;

    while (_info->args.size() < _nArgs)
      _info->args.emplace_back(FuncArg(fmt::format(L"Arg_{}", _info->args.size()).c_str()));

    if (_allowRangeAll)
      for (auto& arg : _info->args)
        arg.allowRange = true;

    if (_info->args.size() > _nArgs)
      XLO_THROW("Too many args for function");

    if ((_info->options & FuncInfo::ASYNC) != 0)
      _info->args.pop_back(); // TODO: hack!!

    return _info;
  }

  std::list<FuncRegistrationMemo>& getFuncRegistryQueue()
  {
    static std::list<FuncRegistrationMemo> theQueue;
    return theQueue;
  }

  XLOIL_EXPORT FuncRegistrationMemo& createRegistrationMemo(const char* entryPoint_, size_t nArgs)
  {
    getFuncRegistryQueue().emplace_back(entryPoint_, nArgs);
    return getFuncRegistryQueue().back();
  }

  std::vector<std::shared_ptr<const FuncSpec>>
    processRegistryQueue(const wchar_t* moduleName)
  {
    std::vector<std::shared_ptr<const FuncSpec>> result;
    auto& queue = getFuncRegistryQueue();
    for (auto f : queue)
      result.emplace_back(make_shared<const StaticSpec>(f.getInfo(), moduleName, f.entryPoint));
    
    queue.clear();
    return result;
  }
}