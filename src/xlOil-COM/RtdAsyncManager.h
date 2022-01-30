#pragma once
#include <xloil/RtdServer.h>
#include <xloil/StringUtils.h>
#include <shared_mutex>

namespace xloil
{
  template<class A, class B>
  struct pair_hash {
    size_t operator()(std::pair<A, B> p) const noexcept
    {
      return xloil::boost_hash_combine(377, p.first, p.second);;
    }
  };

  namespace COM
  {
    struct CellTasks;

    class RtdAsyncManager
    {
    public:
      static RtdAsyncManager& instance();

      std::shared_ptr<const ExcelObj>
        getValue(
          std::shared_ptr<IRtdAsyncTask> task);

      void clear();
    
      using CellAddress = std::pair<unsigned, unsigned>;
      using CellTaskMap = std::unordered_map<
        CellAddress,
        std::shared_ptr<CellTasks>,
        pair_hash<unsigned, unsigned>>;

    private:
      std::shared_ptr<IRtdServer> _rtd;
      CellTaskMap _tasksPerCell;
      mutable std::shared_mutex _mutex;

      RtdAsyncManager();
    };
  }
}