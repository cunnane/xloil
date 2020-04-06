#include "StaticRegister.h"
#include "internal/FuncRegistry.h"
#include "FuncSpec.h"

using std::vector;

namespace xloil
{
  FuncRegistrationMemo::FuncRegistrationMemo(const char* entryPoint_, size_t nArgs)
    : _nArgs(nArgs)
    , entryPoint(entryPoint_)
    , _info(new FuncInfo())
    , _allowRangeAll(false)
  {
    // TODO: why aren't we using the function in Utils?
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> conv;
    _info->name = conv.from_bytes(entryPoint_);
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

  std::vector<RegisteredFuncPtr> processRegistryQueue(const wchar_t* moduleName)
  {
    vector<RegisteredFuncPtr> result;
    auto& queue = getFuncRegistryQueue();
    for (auto f : queue)
    {
      auto spec = std::make_shared<const StaticSpec>(f.getInfo(), moduleName, f.entryPoint);
      auto registeredFunc = registerFunc(spec);
      if (registeredFunc)
        result.emplace_back(registeredFunc);
    }
    queue.clear();
    return result;
  }
}