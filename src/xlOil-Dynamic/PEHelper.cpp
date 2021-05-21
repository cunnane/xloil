#include "PEHelper.h"
#include <cassert>
#include <algorithm>

using xloil::Helpers::Exception;

xloil::DllExportTable::DllExportTable(HMODULE hInstance)
{
  if (!hInstance)
    throw std::runtime_error("DllExportTable: null HINSTANCE");

  // Express image as BYTE pointer as offsets in the PE format
  // are usually in number of bytes
  _imageBase = (BYTE*)hInstance;

  auto* pDosHeader = (PIMAGE_DOS_HEADER)_imageBase;
  if (pDosHeader->e_magic != IMAGE_DOS_SIGNATURE)
    throw Exception("Dll export table: fatal error - bad DOS image");

  auto* pNTHeader = (PIMAGE_NT_HEADERS)(_imageBase + pDosHeader->e_lfanew);
  if (pNTHeader->Signature != IMAGE_NT_SIGNATURE)
    throw Exception("Dll export table: fatal error - bad NT image");
    
  auto optHdr = pNTHeader->OptionalHeader;
  auto& exportEntry = optHdr.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
  auto* pExportDirectory = (PIMAGE_EXPORT_DIRECTORY)(_imageBase + exportEntry.VirtualAddress);
 
  _funcAddresses   = (DWORD*)(_imageBase + pExportDirectory->AddressOfFunctions);
  _namesToOrdinals = (WORD* )(_imageBase + pExportDirectory->AddressOfNameOrdinals);
  _funcNames       = (DWORD*)(_imageBase + pExportDirectory->AddressOfNames);

  _numNames = pExportDirectory->NumberOfNames;
  _numFuncs = pExportDirectory->NumberOfFunctions;
}

int xloil::DllExportTable::findOrdinal(const char* funcName)
{
  // The table function names is lexically ordered. We do need to 
  // to a trick with the comparison functor since the table actually
  // contains offsets from ImageBase rather than strings
  auto found = std::lower_bound(_funcNames, _funcNames + _numNames, 0,
    [base = _imageBase, funcName](DWORD a, DWORD b)
    { 
      return strcmp(
        a > 0 ? (const char*)(base + a) : funcName,
        b > 0 ? (const char*)(base + b) : funcName) < 0;
    }
  );
  if (found != _funcNames + _numNames)
    return (int)_namesToOrdinals[found - _funcNames];
  return -1;
}

#pragma warning(disable: 4302 4311)
bool xloil::DllExportTable::hook(size_t ordinal, void* hook)
{
  if (ordinal >= _numFuncs)
    throw Exception("Function ordinal beyond export table bounds during hook");
  if (hook < _imageBase)
    throw Exception("Hook function must be beyond ImageBase");

  auto* target = _funcAddresses + ordinal;

  DWORD oldProtect;
  if (!VirtualProtect(target, sizeof(DWORD), PAGE_READWRITE, &oldProtect)) 
    return false;

  *target = (BYTE*)hook - _imageBase;

  if (!VirtualProtect(target, sizeof(DWORD), oldProtect, &oldProtect)) 
    return false;

  return true;
}

const char* xloil::DllExportTable::getName(size_t ordinal) const
{
  for (auto i = 0; i < _numNames; ++i)
    if (_namesToOrdinals[i] == ordinal)
      return (const char*)(size_t)_funcNames[i];
  return nullptr;
}

void* xloil::DllExportTable::functionPointer(size_t ordinal) const
{
  return (void*)(_funcAddresses[ordinal] + size_t(_imageBase));
}