#pragma once
#include <xlOilHelpers/Exception.h>
#include <xlOil/WindowsSlim.h>

namespace xloil
{
  
  /// <summary>
  /// Manages hooking into the DLL export table
  /// </summary>
  class DllExportTable
  {
  private:
    DWORD* _funcAddresses;
    WORD* _namesToOrdinals;
    DWORD* _funcNames;
    size_t _numNames;
    size_t _numFuncs;
    BYTE* _imageBase;

  public:
    DllExportTable(HMODULE image);

    /// <summary>
    /// Finds a function's ordinal given its name
    /// </summary>
    /// <param name="funcName"></param>
    /// <returns>The ordinal or -1 if not found</returns>
    int findOrdinal(const char* funcName);

    /// <summary>
    /// Hooks a function at the specified function ordinal, that is, points the the 
    /// export table entry for that function to the hook address. It does not change
    /// the exported function name. The hook function address must be greater than 
    /// the DLL's imagebase.
    /// </summary>
    /// <param name="offset"></param>
    /// <param name="hook"></param>
    /// <returns>true if hook succeeded, else false</returns>
    bool hook(size_t ordinal, void* hook);

    /// <summary>
    /// Returns the exported function name given an ordinal or null pointer if the 
    /// ordinal is out of range or not exported by name
    /// </summary>
    const char* getName(size_t ordinal) const;

    /// <summary>
    /// Returns a pointer to an exported function given its ordinal
    /// </summary>
    /// <param name="ordinal"></param>
    /// <returns></returns>
    void* functionPointer(size_t ordinal) const;
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