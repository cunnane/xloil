#include "PEHelper.h"
#include "xloil/Interface.h"
#include "xloil/Log.h"
#include <algorithm>

xloil::DllExportTable::DllExportTable(HMODULE hInstance)
{
    auto image = (BYTE*)hInstance;
    PIMAGE_DOS_HEADER pDos_hdr = (PIMAGE_DOS_HEADER)image;

    if (pDos_hdr->e_magic != IMAGE_DOS_SIGNATURE)
      OutputDebugStringA("Horrible 1");

    PIMAGE_NT_HEADERS pNt_hdr = (PIMAGE_NT_HEADERS)(image + pDos_hdr->e_lfanew);
    if (pNt_hdr->Signature != IMAGE_NT_SIGNATURE)
      OutputDebugStringA("Horrible 2");
    
    IMAGE_OPTIONAL_HEADER opt_hdr = pNt_hdr->OptionalHeader;
    IMAGE_DATA_DIRECTORY exp_entry = opt_hdr.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
    PIMAGE_EXPORT_DIRECTORY pExp_dir = (PIMAGE_EXPORT_DIRECTORY)(image + exp_entry.VirtualAddress); //Get a pointer to the export directory

    func_table = (DWORD*)(image + pExp_dir->AddressOfFunctions);
    ord_table = (WORD*)(image + pExp_dir->AddressOfNameOrdinals);
    name_table = (DWORD*)(image + pExp_dir->AddressOfNames);
    numberOfNames = pExp_dir->NumberOfNames;
    imageBase = image;

    if (pExp_dir->NumberOfFunctions != pExp_dir->NumberOfNames)
      XLO_THROW("Dll is exporting functions by ordinal, we don't currently support this");
}

size_t xloil::DllExportTable::findOffset(const char * funcName)
{
  // TODO: lowerBound - the name table is sorted!
  /*auto found = std::lower_bound(name_table, name_table + numberOfNames, funcName,
    [this](DWORD a, const char* b) { return strcmp() < 0; }
  );*/
  for (auto i = 0; i < numberOfNames; ++i)
  {
    if (strcmp((char*)imageBase + name_table[i], funcName) == 0)
      return i;
  }
  abort();
  return -1;
}

#pragma warning(disable: 4302 4311)
bool xloil::DllExportTable::hook(size_t offset, void * hook)
{
  if (offset >= numberOfNames)
    throw std::exception();
  auto target = func_table + ord_table[offset];
  DWORD oldProtect;
  if (!VirtualProtect(target, sizeof(DWORD), PAGE_READWRITE, &oldProtect)) return false;
  *target = (DWORD)hook - DWORD(imageBase);
  if (!VirtualProtect(target, sizeof(DWORD), oldProtect, &oldProtect)) return false;;
  return true;
}

