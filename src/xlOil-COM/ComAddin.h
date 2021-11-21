#pragma once
#include <memory>
namespace xloil
{
  class IComAddin;

  namespace COM
  {
    std::shared_ptr<IComAddin> createComAddin(const wchar_t* name, const wchar_t* description);
  }
}