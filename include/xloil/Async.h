#pragma once
#include "ExportMacro.h"
#include <xloil/ExcelObj.h>
#include <xloil/Events.h>
#include <memory>
namespace xloil { class ExcelObj; }
namespace xloil
{
  /// <summary>
  /// Wrapper for xlAsyncReturn which takes ExcelObj arguments. Used
  /// to return values from native async functions.
  /// </summary>
  XLOIL_EXPORT void 
    asyncReturn(
      const ExcelObj& asyncHandle, const ExcelObj& value);

  struct AsyncHandle : public ExcelObj
  {
    template<class... Args>
    void returnValue(Args&&... args) const
    {
      asyncReturn(*this, ExcelObj(std::forward<Args>(args)...));
    }
    void returnValue(const ExcelObj& value)
    {
      asyncReturn(*this, value);
    }
  };

  XLOIL_EXPORT bool yieldAndCheckIfEscPressed();

  class AsyncHelper
  {
    std::shared_ptr<const void> _eventHandler;
    ExcelObj _asyncHandle;

  public:
    AsyncHelper(const ExcelObj& asyncHandle)
      : _asyncHandle(asyncHandle)
    {
      _eventHandler = xloil::Event::CalcCancelled().bind(
          [self = this]() { self->cancel(); });
    }
    virtual ~AsyncHelper()
    {
      _eventHandler.reset();
    }
    virtual void cancel()
    {
      _eventHandler.reset();
    }
    void result(const ExcelObj& value)
    {
      asyncReturn(_asyncHandle, value);
    }
  };
}