#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/ExportMacro.h>
#include <memory>
#include <future>

namespace xloil
{
  /// <summary>
  /// An instance of this interface will be passed to an RtdTask.
  /// The <see cref="RtdTask"/> should indicate new data with publish.
  /// </summary>
  struct IRtdPublish
  {
    /// <summary>
    /// Passes a value to underlying Rtd server, which will trigger an
    /// update in Excel.
    /// </summary>
    /// <param name=""></param>
    virtual bool publish(ExcelObj&& value) noexcept = 0;
  };

  /// <summary>
  /// An IRtdTask is a `std::packaged_task`-like object which can start and stop 
  /// a background publishing task.  It returns results through a
  /// <see cref="IRtdPublish"/> rather than directly.  
  /// 
  /// If awaiting multiple futures and cancellation were supported in the STL's
  /// async library, this interface could be factored out.
  /// </summary>
  struct IRtdTask
  {
    virtual ~IRtdTask() {}

    /// <summary>
    /// Start task, writing updates to the giver publisher
    /// </summary>
    /// <param name=""></param>
    virtual void start(IRtdPublish&) = 0;

    /// <summary>
    /// Return true if all child tasks have been cleanly shut down and this object
    /// can be destroyed
    /// </summary>
    /// <returns></returns>
    virtual bool done() = 0;

    /// <summary>
    /// Wait for pending completion of child tasks
    /// </summary>
    virtual void wait() = 0;

    /// <summary>
    /// Request publication ceases and child tasks be shut down
    /// </summary>
    virtual void cancel() = 0;
  };

  /// <summary>
  /// Associates a topic string with a publishing task (typically a 
  /// <see cref="IRtdTask"/>) and manages publication of results via
  /// <see cref="IRtdServer::publish"/>. This interface does not extend
  /// <see cref="IRtdPublish">, which publishes the actual values, rather
  /// it defines the management of the task in the context of the RTD Server.
  /// </summary>
  struct IRtdPublisher
  {
    /// <summary>
    /// Called when a worksheet function subscribes to the topic.
    /// </summary>
    /// <param name="numSubscribers">current number of subscibers (including this one)</param>
    virtual void connect(size_t numSubscribers) = 0;

    /// <summary>
    /// Called when a worksheet function unsubscribes to the topic. e.g. because
    /// a formula has been changed or deleted.
    /// </summary>
    /// <param name="numSubscribers"></param>
    /// <returns>return true if you want the RtdServer to destroy this topic</returns>
    virtual bool disconnect(size_t numSubscribers) = 0;

    /// <summary>
    /// Called by the RtdServer to tell the topic to stop all child workers
    /// </summary>
    virtual void stop() = 0;

    /// <summary>
    /// Return true if all child workers have cleanly shutdown and the object
    /// can be destroyed
    /// </summary>
    /// <returns></returns>
    virtual bool done() const = 0;

    /// <summary>
    /// The name of the topic
    /// </summary>
    /// <returns></returns>
    virtual const wchar_t* topic() const = 0;
  };
;

  /// <summary>
  /// An IRtdAsyncTask packages a job for a worksheet functions which uses 
  /// RTD support to allow background execution. The function which initiates
  /// the job will be called by Excel again to collect the result. Since the 
  /// worksheet function is stateless it cannot tell if it should start a new 
  /// task or collect a result. Hence the task object must support the '==' 
  /// operator so xlOil can  compare the task to be started to any that have 
  /// pending results for the calling cell.  This comparison carries some 
  /// overhead, so the RTD async mechanism should only be used when this 
  /// overhead is small relative to the function execution time.
  /// </summary>
  struct IRtdAsyncTask : public IRtdTask
  {
    virtual bool operator==(const IRtdAsyncTask& that) const = 0;
  };

  /// <summary>
  /// Extends <see cref="IRtdPublish"/> to provide a cancellation token. C++
  /// does not currently support cancellable futures, so cancellation is 
  /// usually via periodic checking of a bool.
  /// </summary>
  struct RtdNotifier
  {
    RtdNotifier(
      IRtdPublish& publisher,
      const std::atomic<bool>& cancelFlag)
      : _publisher(publisher)
      , _cancel(cancelFlag)
    {}

    /// <summary>
    /// This flag should be periodically checked to ensure cancellation is 
    /// responsive. If it is true, the producer should immediately shut down
    /// </summary>
    /// <returns></returns>
    bool cancelled() const
    {
      return _cancel;
    }

    bool publish(ExcelObj&& value) const noexcept
    {
      return _cancel 
        ? false 
        : _publisher.publish(std::forward<ExcelObj>(value));
    }

  private:
    IRtdPublish& _publisher;
    const std::atomic<bool>& _cancel;
  };


  /// <summary>
  /// Concrete implemenation of <see cref="IRtdTask"/> which implements the
  /// interface using a <code>std::future<void></code>.
  /// </summary>
  template <class TBase>
  class RtdTaskBase : public TBase
  {
  public:
    virtual std::future<void> operator()(RtdNotifier notify) = 0;

    virtual void start(IRtdPublish& n) override
    {
      // The producer may be stopped and restarted, so wait for any prior
      // future to exit
      wait();
      _cancel = false;
      _future = (*this)(RtdNotifier(n, _cancel));
    }
    bool done() override
    {
      return !_future.valid()
        || _future.wait_for(std::chrono::seconds(0)) == std::future_status::ready;
    }
    void wait() override
    {
      if (_future.valid())
        _future.wait();
    }
    void cancel() override
    {
      _cancel = true;
    }

  private:
    std::future<void> _future;
    std::atomic<bool> _cancel = false;
  };

  /// <summary>
  /// Wraps a <code>std::function</code> to make an IRtdTask. The function 
  /// should take an RtdNotifier and return a <code>std::future<void> </code>.
  /// The future is just a synchronisation object - returned values should be
  /// published through the RtdNotifier. The cancel flag in the notifier should be
  /// periodically checked.
  /// </summary>
  class RtdTask : public RtdTaskBase<IRtdTask>
  {
  public:
    using Prototype = std::function<std::future<void>(RtdNotifier)>;

    RtdTask(const Prototype& func)
     : _func(func)
    {}

    std::future<void> operator()(RtdNotifier notify) override
    {
      return _func(notify);
    }

  private:
    Prototype _func;
  };

  /// <summary>
  /// A base class for Rtd async tasks. 
  /// </summary>
  struct RtdAsyncTask : public RtdTaskBase<IRtdAsyncTask>
  {};

  class IRtdServer;

  /// <summary>
  /// Concrete implementation of <see cref="IRtdTopc"/> which can be overriden to 
  /// hook the virtual methods. 
  /// </summary>
  class XLOIL_EXPORT RtdPublisher : public IRtdPublisher, public IRtdPublish
  {
  public:
    RtdPublisher(
      const wchar_t* topic,
      IRtdServer& mgr,
      const std::shared_ptr<IRtdTask>& task);
    virtual ~RtdPublisher();
    virtual void connect(size_t numSubscribers) override;
    virtual bool disconnect(size_t numSubscribers) override;
    virtual void stop() override;
    virtual bool done() const override;
    virtual const wchar_t* topic() const override;
    virtual bool publish(ExcelObj&& value) noexcept override;
    const std::shared_ptr<IRtdTask>& task() const { return _task; }

  protected:
    std::shared_ptr<IRtdTask> _task;
    IRtdServer& _mgr;
    std::wstring _topic;
  };
  

  /// <summary>
  /// An IRtdServer is a wrapper around an RTD COM Server. An RTD Server is a
  /// producer/consumer queue which can trigger recalculations in cells marked
  /// marked as RTD subscribers (consumers).  Excel will recalculate the entire
  /// cell containing a subscriber, it cannot distinguish between multiple 
  /// functions in a single cell.  
  /// 
  /// An RTD producer can be started anywhere including in another cell, or 
  /// even the same cell as the consumer. The latter allows execution of 
  /// functions asynchronously without the drawback of Excel's asynchronous 
  /// UDF support, which is that async functions are cancelled if the user 
  /// interacts with the sheet.
  /// 
  /// RTD producers and subscribers find each other using a topic string. The
  /// producer and subscribers can be registered in either order.
  /// </summary>
  class IRtdServer
  {
  public:
    /// <summary>
    /// Starts a producer embedded in an <see cref="RtdTask"/>
    /// </summary>
    /// <param name="task"></param>
    /// <param name="topic">The topic key which consumers can use to 
    /// locate this producer.
    /// </param>
    /// <param name="persistent">If false, the producer will be cancelled
    /// when its subscriber count reaches zero for a second time (the first
    /// time being at creation). False is the natural choice for a producer
    /// and consumer started in the same cell
    /// </param>
    virtual void start(
      const std::shared_ptr<IRtdPublisher>& topic) = 0;

    /// <summary>
    /// Convenience function which wraps a given `std::function` object in a 
    /// <see cref="IRtdPublisher"/>, then calls <see cref="IRtdServer::start"/>
    /// </summary>
    /// 
    /// <param name="topic"> Topic name to be associated with the output of the 
    ///   function
    /// </param>
    /// <param name="func"> The function must have a specific form, see 
    ///   <see cref="RtdTask"/> for details
    /// </param>
    void start(
      const wchar_t* topic,
      const RtdTask::Prototype& func)
    {
      start(
        std::make_shared<RtdPublisher>(
          topic, *this, std::make_shared<RtdTask>(func)));
    }
    /// <summary>
    /// Subscribes to a the specified topic. If no publisher for the topic
    /// currently exists, the subscription will be held open pending
    /// one being created. This calls Excel's RTD function, which means the
    /// calling cell will be recalculated every time a new value is published.
    /// 
    /// Calling this method when not in a worksheet function invoked by Excel
    /// is undefined behaviour.
    /// </summary>
    /// <param name="topic"></param>
    /// <returns>The ExcelObj currently being published by the publisher
    ///   or an empty ptr if no publisher exists.
    /// </returns>
    virtual std::shared_ptr<const ExcelObj> 
      subscribe(
        const wchar_t* topic) = 0;

    /// <summary>
    /// Looks up a value for a specified topic, but does not subscribe.
    /// If there is no publisher for the topic, the returned pointer will
    /// be null. If there is no published value, it will point to N/A.
    /// This does not call Excel's RTD function.
    /// </summary>
    /// <param name="topic"></param>
    /// <returns></returns>
    virtual std::shared_ptr<const ExcelObj>
      peek(
        const wchar_t* topic) = 0;

    /// <summary>
    /// Force publication of the specified value by the producer of the 
    /// given topic. If it is actively publishing values, the producer 
    /// will override this setting at the next update.
    /// </summary>
    /// <param name="topic"></param>
    /// <param name="value"></param>
    /// <returns>True if the producer was found and the value was set</returns>
    virtual bool 
      publish(
        const wchar_t* topic,
        ExcelObj&& value) = 0;

    /// <summary>
    /// Drops the producer for a topic by calling RtdPublisher::stop, then waits
    /// for it to complete and publishes #N/A to all subscribers.
    /// </summary>
    /// <param name="topic"></param>
    /// <returns></returns>
    virtual bool 
      drop(const wchar_t* topic) = 0;

    /// <summary>
    /// Drop all topics, stop their tasks and wait for completion
    /// </summary>
    virtual void 
      clear() = 0;

    /// <summary>
    /// Get the ProgId associated with the underlying COM server
    /// </summary>
    /// <returns></returns>
    virtual const wchar_t* progId() const noexcept = 0;
  };

  /// <summary>
  /// Can be called from a worksheet function to run the given task asynchronously 
  /// in the context of Excel using the RTD machinery. This function does not 
  /// ensure the task will be run on another thread or process - the task must 
  /// ensure it is not blocking the calling (main) thread.
  /// </summary>
  XLOIL_EXPORT std::shared_ptr<ExcelObj> rtdAsync(
    const std::shared_ptr<IRtdAsyncTask>& task);

  void rtdAsyncServerClear();

  /// <summary>
  /// Creates a new Rtd Manager.  Optionally wraps the an Excel::IRtdServer COM
  /// object specified with progId and clsid. The necessary registry keys to
  /// access this COM object will be created if required.
  /// </summary>
  XLOIL_EXPORT std::shared_ptr<IRtdServer>
    newRtdServer(
      const wchar_t* progId = nullptr,
      const wchar_t* clsid = nullptr);
}