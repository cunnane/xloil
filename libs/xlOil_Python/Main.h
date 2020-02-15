#pragma once
#include <xloil/Interface.h>
namespace xloil
{
  namespace Python
  {
    extern Core* theCore;
    Event<void(void), VoidCollector>& Event_PyBye();
  }
}