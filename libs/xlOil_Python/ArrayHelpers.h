#pragma once
#include "PyHelpers.h"
#include "Cache.h"

namespace xloil
{
  namespace Python
  {
    inline void accumulateObjectStringLength(PyObject* p, size_t& strLength)
    {
      if (PyUnicode_Check(p))
        strLength += PyUnicode_GetLength(p);
      else if (!PyFloat_Check(p) && !PyLong_Check(p) && !PyBool_Check(p))
        strLength += CACHE_KEY_MAX_LEN;
    }
  }
}