#pragma once

namespace Office { struct ICTPFactory; }

namespace xloil
{
  class ICustomTaskPane;

  namespace COM
  {
    ICustomTaskPane* createCustomTaskPane(
      Office::ICTPFactory& ctpFactory, 
      const wchar_t* name,
      const wchar_t* progId = nullptr);
  }
}