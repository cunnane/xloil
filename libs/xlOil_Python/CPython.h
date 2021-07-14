#pragma once
// Horrible hack to allow our debug build to link with release python lib and so avoid building debug python
// Can remove this for Python >= 3.8 (still seems to be a problem wth 3.9)
// https://stackoverflow.com/questions/17028576/using-python-3-3-in-c-python33-d-lib-not-found
#ifdef _DEBUG
#  define XLO_PY_HACK
#endif
#undef _DEBUG
#define HAVE_SNPRINTF
#include <Python.h>
#ifdef XLO_PY_HACK
#  define _DEBUG
#endif
