#include <xloil/Register/FuncRegistry.h>
#include <xloil/StaticRegister.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/Preprocessor.h>
#include <xloil/Log.h>
#include <spdlog/sinks/basic_file_sink.h>

namespace xloil
{
  XLO_FUNC_START(xloLog())
  {
    spdlog::default_logger()->flush();
    // TODO: better to add the log file name to the addin context?
    for (auto& sink : spdlog::default_logger()->sinks())
    {
      auto p = dynamic_cast<spdlog::sinks::basic_file_sink_mt*>(sink.get());
      if (p)
        return returnValue(p->filename());
    }
    return returnValue(CellError::NA);
  }
  XLO_FUNC_END(xloLog).threadsafe()
    .help(L"Flushes the log and returns the location of the log file");
}