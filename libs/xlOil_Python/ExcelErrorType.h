#pragma once

struct _typeobject;

namespace xloil {
  namespace Python {
    /// <summary>
    /// Type object correponding to the bound xloil::CellError
    /// </summary>
    extern _typeobject* pyExcelErrorType;
  }
}