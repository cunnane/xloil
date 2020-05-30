#pragma once
#include <xloil/ExcelObj.h>
#include <memory>
#include <future>

namespace xloil
{
  namespace COM
  {
    struct IRtdNotify
    {
      virtual void setValue(ExcelObj&&) = 0;
      virtual bool isCancelled() const = 0;
    };

    using RtdTask = std::function<std::future<void>(IRtdNotify&)>;

    class IRtdManager
    {
    public:
      virtual ExcelObj start(
        const wchar_t* topic, 
        const RtdTask& task) = 0;

      virtual bool getValue(
        const wchar_t* topic,
        ExcelObj& val) = 0;

      virtual bool setValue(
        const wchar_t* jobRef, 
        ExcelObj&& value) = 0;
    };

    std::shared_ptr<IRtdManager> 
      newRtdManager(
        const wchar_t* progId = nullptr, 
        const wchar_t* clsid = nullptr);

    class RtdConnection
    {
    public:
      RtdConnection(IRtdManager& mgr, std::wstring&& topic);
      bool hasValue() const;
      ExcelObj&& value();
      ExcelObj start(const RtdTask& task);

      ExcelObj run(const RtdTask& task)
      {
        return hasValue() ? value() : start(task);
      }

    private:
      std::wstring _topic;
      IRtdManager& _mgr;
      ExcelObj _value;
    };

    /// <summary>
    ///  
    /// <example>
    /// <code>
    ///   auto p = rtdConnect();
    ///   return p.hasValue() 
    ///       ? p.value() 
    ///       : p.start([](notify) { notify.setValue(ExcelObj(1)); } );
    /// </code>
    /// </example>
    /// </summary>
    RtdConnection rtdConnect(
      IRtdManager* mgr = nullptr,
      const wchar_t* topic = nullptr);
  }
}