=======================
xlOil C++ RTD functions
=======================

In Excel, RTD (real time data) functions are able to return values independently of Excel's 
calculation cycle. The classic example of this is a stock ticker with live prices. This is contrasted 
with  Excel's native async support in :any:`concepts-rtd-async`.


Example: Subscription
---------------------

xlOil wraps the RTD server implementation to make subscription and publication straightfoward as in the 
following example:

.. highlight:: c++

::

    #include <xloil/xlOil.h>
    using namespace xloil;

    IRtdServer& myRtdServer()
    {
      static shared_ptr<IRtdServer> ptr = newRtdServer();
      static auto backgroundTask = std::async([]()
        {
          // Somehow fetch live prices here
          ptr->publish("SP500", ExcelObj(123));
          ptr->publish("FTSE100", ExcelObj(456));
        });
      return *ptr;
    }

    XLO_FUNC_START(
      getStock(const ExcelObj& ticker)
    )
    {
      auto value = myRtdServer().subscribe(tag.toString().c_str());
      return returnValue(value
        ? *value
        : Const::Error(CellError::NA));
    }
    XLO_FUNC_END(getStock);


Example: Background task
------------------------

The RTD mechanism can be used to run slow tasks like data fetches or calculations in the background whilst
the user continues to interact with Excel.  The following demonstrates a simple counter which increments
stepwise.

.. highlight:: c++

::

    #include <xloil/xlOil.h>
    using namespace xloil;
    
    struct Counter : public RtdAsyncTask
    {
      Counter(int iStep) : _iStep(iStep) {}

      int _iStep;

      // This function arranges for the work to be done in the background, but must
      // exit when the `RtdNotifier` requests it
      std::future<void> operator()(RtdNotifier notify) override
      {
        return std::async([=, step = _iStep]()
        {
          int _count = 0;
          while (!notify.cancelled())
          {
            notify.publish(ExcelObj(_count));
            std::this_thread::sleep_for(std::chrono::seconds(2));
            _count += step;
          }
        });
      }

      // When the cell which generated the task is recalculated, xlOil cannot tell
      // whether Excel is calling it with new inputs or RTD triggered the call to get
      // a result. It uses this equality operator to determine if it already has a 
      // running task for the given set of inputs
      bool operator==(const IRtdAsyncTask& that_) const override
      {
        const auto* that = dynamic_cast<const Counter*>(&that_);
        if (!that)
          return false;
        return _iStep == that->_iStep;
      }
    };

    XLO_FUNC_START(
      rtdCounter(const ExcelObj& step)
    )
    {
      auto value = rtdAsync(
        std::make_shared<Counter>(step.toInt(1)));
      return returnValue(value ? *value : CellError::NA);
    }
    XLO_FUNC_END(rtdCounter);