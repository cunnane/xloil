#pragma once

#ifdef XLOIL_EXPORTS
#define XLOIL_EXPORT __declspec(dllexport)
#else
#ifdef XLOIL_STATIC_LIB
#define XLOIL_EXPORT 
#else
#define XLOIL_EXPORT __declspec(dllimport)
#endif
#endif

#define XLO_ENTRY_POINT(ret) extern "C" __declspec(dllexport) ret __stdcall 

// Ensure entry points from XllEntryPoint are exposed
#ifdef XLOIL_STATIC_LIB
#pragma comment(linker, "/include:DllMain")
#pragma comment(linker, "/include:xlAutoOpen")
#pragma comment(linker, "/include:xlAutoClose")
#pragma comment(linker, "/include:xlAutoFree12")
#pragma comment(linker, "/include:xlHandleCalculationCancelled")
#endif