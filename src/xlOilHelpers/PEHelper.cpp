#include "PEHelper.h"
#include <xloil/Throw.h>
#include <algorithm>

xloil::DllExportTable::DllExportTable(HMODULE hInstance)
{
  // Express image as BYTE pointer as offsets in the PE format
  // are usually in number of bytes
  imageBase = (BYTE*)hInstance;

  auto* pDosHeader = (PIMAGE_DOS_HEADER)imageBase;
  if (pDosHeader->e_magic != IMAGE_DOS_SIGNATURE)
    throw std::runtime_error("Dll export table: fatal error - bad DOS image");

  auto* pNTHeader = (PIMAGE_NT_HEADERS)(imageBase + pDosHeader->e_lfanew);
  if (pNTHeader->Signature != IMAGE_NT_SIGNATURE)
    throw std::runtime_error("Dll export table: fatal error - bad NT image");
    
  auto opt_hdr = pNTHeader->OptionalHeader;
  auto& exp_entry = opt_hdr.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
  auto* pExportDirectory = (PIMAGE_EXPORT_DIRECTORY)(imageBase + exp_entry.VirtualAddress);
 

  func_table = (DWORD*)(imageBase + pExportDirectory->AddressOfFunctions);
  ord_table = (WORD*)(imageBase + pExportDirectory->AddressOfNameOrdinals);
  name_table = (DWORD*)(imageBase + pExportDirectory->AddressOfNames);
  numberOfNames = pExportDirectory->NumberOfNames;

  if (pExportDirectory->NumberOfFunctions != pExportDirectory->NumberOfNames)
    throw std::runtime_error("Dll export table: ordinal export detected - not currently supported");
}

int xloil::DllExportTable::findOffset(const char* funcName)
{
  // TODO: lowerBound - the name table is sorted!
  /*auto found = std::lower_bound(name_table, name_table + numberOfNames, funcName,
    [this](DWORD a, const char* b) { return strcmp() < 0; }
  );*/
  for (size_t i = 0; i < numberOfNames; ++i)
  {
    if (strcmp((char*)imageBase + name_table[i], funcName) == 0)
      return (int)i;
  }
  return -1;
}

#pragma warning(disable: 4302 4311)
bool xloil::DllExportTable::hook(size_t offset, void * hook)
{
  if (offset >= numberOfNames)
    throw std::runtime_error("Function offset out of bounds of export table during hook");
  auto target = func_table + ord_table[offset];
  DWORD oldProtect;
  if (!VirtualProtect(target, sizeof(DWORD), PAGE_READWRITE, &oldProtect)) return false;
  *target = (DWORD)hook - DWORD(imageBase);
  if (!VirtualProtect(target, sizeof(DWORD), oldProtect, &oldProtect)) return false;;
  return true;
}

