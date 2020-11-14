#pragma once
#include <xlOilHelpers/Exception.h>
#include <xlOil/WindowsSlim.h>

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

    int findOffset(const char* funcName);

    DWORD* getAddress(size_t offset) const;

    /// Hooks a function at the specified function number.  Currently the function address must be 
    /// greater than the DLL's imagebase.
    bool hook(size_t offset, void* hook);

    const char* getName(size_t offset) const
    {
      if (offset >= numberOfNames)
        throw Helpers::Exception("Function offset out of bounds of export table");
      return (const char*)(imageBase + name_table[offset]);
    }
  };


  // With Win32 function C function names are decorated. It no longer 
  // seemed like a good idea with x64.
  inline std::string 
    decorateCFunction(const char* name, const size_t numPtrArgs)
  {
#ifdef _WIN64
    (void)numPtrArgs;
    return std::string(name);
#else
    return formatStr("_%s@%d", name, sizeof(void*) * numPtrArgs);
#endif // _WIN64
  }
}