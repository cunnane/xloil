#pragma once

#ifdef XLOIL_EXPORTS
#define XLOIL_EXPORT __declspec(dllexport)
#else
#define XLOIL_EXPORT __declspec(dllimport)
#endif

#define XLO_ENTRY_POINT(ret) extern "C" __declspec(dllexport) ret __stdcall 
