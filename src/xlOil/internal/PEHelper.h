#pragma once
#include <xloil/WindowsSlim.h>

namespace xloil
{
  /// Hides the casting around inspecting and hooking the DLL's export address table (EAT)
  class DllExportTable
  {
  private:
    DWORD* func_table;
    WORD* ord_table;
    DWORD * name_table;
    size_t numberOfNames;
    BYTE* imageBase;

  public:
    DllExportTable(HMODULE image);

    size_t findOffset(const char* funcName);

    DWORD* getAddress(size_t offset) const;

    /// Hooks a function at the specified function number.  Currently the function address must be 
    /// greater than the DLL's imagebase.
    bool hook(size_t offset, void* hook);

    const char* getName(size_t offset) const
    {
      return (const char*)(imageBase + name_table[offset]);
    }
  };
}