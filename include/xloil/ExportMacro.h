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

#define XLO_ENTRY_POINT(return_value) extern "C" __declspec(dllexport) return_value __stdcall 
