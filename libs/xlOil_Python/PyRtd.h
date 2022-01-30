#pragma once
#include "PyCore.h"
#include <xlOil/RtdServer.h>
#include <pybind11/pybind11.h>

namespace xloil
{
  class CallerInfo;

  namespace Python
  {
    class RtdReturn
    {
    public:
      RtdReturn(
        IRtdPublish& notify,
        const std::shared_ptr<const IPyToExcel>& returnConverter,
        const CallerInfo& caller);
      ~RtdReturn();
      void set_task(const pybind11::object& task);
      void set_result(const pybind11::object& value) const;
      void set_done();
      void cancel();
      bool done() noexcept;
      void wait() noexcept;
      const CallerInfo& caller() const noexcept;

    private:
      IRtdPublish& _notify;
      std::shared_ptr<const IPyToExcel> _returnConverter;
      pybind11::object _task;
      std::atomic<bool> _running;
      const CallerInfo& _caller;
    };
  }
}