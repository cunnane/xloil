#pragma once
#include <xloil/ExcelObj.h>
#include <xlOil/ExportMacro.h>
#include <memory>
#include <future>

namespace xloil
{
  /// <summary>
  /// An instance of this interface will be passed to an RtdTask.
  /// The <see cref="RtdTask"/> should poll isCancelled() and indicate 
  /// new data with publish.
  /// </summary>
  struct IRtdNotify
  {
    /// <summary>
    /// Passes a value to underlying Rtd server, which will trigger an
    /// update in Excel.
    /// </summary>
    /// <param name=""></param>
    virtual void publish(ExcelObj&&) = 0;

    /// <summary>
    /// If this returns true, the enclosing future should exit.
    /// </summary>
    /// <returns></returns>
    virtual bool isCancelled() const = 0;
  };

  /// <summary>
  /// A function object which should create a future. The code behind the future
  /// should return data via <see cref="IRtdNotify"/> rather than the return 
  /// statement.  It may run indefinitely, but should poll for cancellation
  /// via <see cref="IRtdNotify::isCancelled"/>
  /// </summary>
  using RtdTask = std::function<std::future<void>(IRtdNotify&)>;

  /// <summary>
  /// An IRtdManager is a wrapper around an internal RTD Server. An RTD Server 
  /// is a producer/consumer queue which can trigger recalculations in 
  /// cells marked as RTD consumers.  An RTD producer can be started anywhere
  /// including in another cell, or even the same cell as the consumer. The 
  /// latter allows execution of functions asynchronously without the drawback
  /// of Excel's asynchronous UDF support, which is that async functions are 
  /// cancelled if the user interacts with the sheet.
  /// 
  /// RTD producers and consumers find each other using a topic string. The
  /// producer and consumer can be registered in either order.
  /// </summary>
  class IRtdManager
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
      const RtdTask& task,
      const wchar_t* topic,
      bool persistent = false) = 0;

    /// <summary>
    /// Subscribes to a producer with the specified topic. If no producer
    /// for the topic currently exists, the subscription will be held open
    /// pending a producer. This calls Excel's RTD function, which means the
    /// calling cell will be recalculated every time the producer published
    /// a new value.
    /// </summary>
    /// <param name="topic"></param>
    /// <returns>The ExcelObj currently being published by the producer
    /// or null if no producer exists.
    /// </returns>
    virtual std::shared_ptr<const ExcelObj> subscribe(
      const wchar_t* topic) = 0;

    /// <summary>
    /// Looks up a value for a specified producer if it exists without 
    /// subscribing. Does not call Excel's RTD function.
    /// </summary>
    /// <param name="topic"></param>
    /// <returns></returns>
    virtual std::shared_ptr<const ExcelObj> peek(
      const wchar_t* topic) = 0;

    /// <summary>
    /// Force publication of the specified value by the producer of the 
    /// given topic. If it is actively publishing values, the producer 
    /// will override this setting at the next update.
    /// </summary>
    /// <param name="topic"></param>
    /// <param name="value"></param>
    /// <returns>True if the producer was found and the value was set</returns>
    virtual bool publish(
      const wchar_t* topic,
      ExcelObj&& value) = 0;
  };


  /// <summary>
  /// Creates a new Rtd Manager.  Optionally wraps the an Excel::IRtdServer COM
  /// object specified with progId and clsid. The necessary registry keys to
  /// access this COM object will be created if required.
  /// </summary>
  XLOIL_EXPORT std::shared_ptr<IRtdManager>
    newRtdManager(
      const wchar_t* progId = nullptr,
      const wchar_t* clsid = nullptr);


  /// <summary>
  /// RtdConnection is a convenience wrapper around <see cref="IRtdManager"/>.
  /// It is created by <see cref="rtdConnect"/>.
  /// </summary>
  class XLOIL_EXPORT RtdConnection
  {
  public:
    RtdConnection(IRtdManager& mgr, std::wstring&& topic);
    /// <summary>
    /// Returns true if there is a value available for the connected topic
    /// </summary>
    bool 
      hasValue() const;
    /// <summary>
    /// Returns the value for the connected topic or # NULL! if unavailable
    /// </summary>
    const ExcelObj& 
      value();
    /// <summary>
    /// Starts the specified task as a producer for the connected topic.
    /// </summary>
    /// <returns>The published value (which may be # GETTING_DATA)</returns>
    const ExcelObj& 
      start(const RtdTask& task);
    /// <summary>
    /// Returns the connected value if available, otherwise starts the
    /// specified task and returns the value (which may be # GETTING_DATA)
    /// </summary>
    const ExcelObj& 
      run(const RtdTask& task)
    {
      return hasValue()
        ? value()
        : start(task);
    }
  private:
    std::wstring _topic;
    IRtdManager& _mgr;
    const ExcelObj* _value;
  };

  /// <summary>
  /// Connects to the Core RtdManager or the one specified, returning an
  /// RtdConnection. The Core RtdManager is only accessible through this 
  /// function. 
  /// <example>
  /// <code>
  ///   auto p = rtdConnect();
  ///   return p.hasValue() 
  ///       ? p.value() 
  ///       : p.start([](notify) { notify.publish(ExcelObj(1)); } );
  /// </code>
  /// </example>
  /// </summary>
  /// <param name="mgr">Omit / null to use the Core RtdManager</param>
  /// <param name="topic">Omit / null to use the current cell address
  /// as the topic
  /// </param>
  XLOIL_EXPORT RtdConnection rtdConnect(
    IRtdManager* mgr = nullptr,
    const wchar_t* topic = nullptr);
}