#pragma once
// Horrible hack to allow our debug build to link with release python lib and so avoid building debug python
// https://stackoverflow.com/questions/17028576/

// Include these before Python.h does or our trick with the _DEBUG
// macro will backfire
#  include <stdlib.h>
#  include <stdio.h>
#  include <errno.h>
#  include <string.h>   

#ifdef _DEBUG
#  define XLO_PY_HACK
#endif
#undef _DEBUG
#include <Python.h>
#ifdef XLO_PY_HACK
#  define _DEBUG
#endif