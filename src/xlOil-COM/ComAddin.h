#pragma once
#include <memory>
#include <string>
#include <map>
#include <functional>
namespace xloil
{
  class IComAddin;

  namespace COM
  {
    std::shared_ptr<IComAddin> createComAddin(
      const wchar_t* name, const wchar_t* description);
  }
}