#pragma once
// Horrible hack to allow our debug build to link with release python lib and so avoid building debug python
// Can remove this for Python >= 3.10 as we have the debug libs
// https://stackoverflow.com/questions/17028576/


#if PY_VERSION_HEX < 0x03100000
#ifdef _DEBUG
#  define XLO_PY_HACK
#endif
#undef _DEBUG
#include <Python.h>
#ifdef XLO_PY_HACK
#  define _DEBUG
#endif
#else
#include <Python.h>
#endif