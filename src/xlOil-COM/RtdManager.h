#pragma once
#include <xloil/RtdServer.h>

struct tagSAFEARRAY;
using SAFEARRAY = tagSAFEARRAY;

namespace xloil
{
  namespace COM
  {
    struct IRtdServerWorker
    {
      void start(std::function<void()>&& updateNotify);
      void connect(long topicId, std::wstring&& topic);
      void disconnect(long topicId);
      SAFEARRAY* getUpdates();
      void quit();
    };

    template <class TValue>
    struct IRtdPublishManager
    {
      void update(const wchar_t* topic, const std::shared_ptr<TValue>& value);
      void addPublisher(const std::shared_ptr<IRtdPublisher>& job);
      bool dropPublisher(const wchar_t* topic);
      bool value(const wchar_t* topic, std::shared_ptr<const TValue>& val) const;
      void quit();
      void join();
    };

    std::shared_ptr<IRtdServer> newRtdServer(
      const wchar_t* progId, const wchar_t* clsid);
  }
}
