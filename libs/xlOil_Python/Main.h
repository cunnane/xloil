#pragma once
#include <xloil/Interface.h>
#include <future>

namespace xloil
{
  namespace Python
  {
    /// <summary>
    /// The addin context of the main xloil.dll
    /// </summary>
    extern AddinContext* theCoreContext;

    /// <summary>
    /// The current context is set to reflect the addin whose
    /// settings are being processed. It is then switched back
    /// to the core context.
    /// </summary>
    extern AddinContext* theCurrentContext;


    template<typename F>
    inline auto runPython(F&& f) -> std::future<decltype(f())> 
    {
      auto pck = std::make_shared<std::packaged_task<decltype(f())()>>(std::forward<F>(f));
      auto _f = std::function<void(int /*id*/)>(
        [pck](int id) {
          (*pck)();
        });
      runPython(std::move(_f));
      return pck->get_future();
    }

    void runPython(std::function<void(int)>&& task);
  }
}