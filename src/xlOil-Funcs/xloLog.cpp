#include <xloil/StaticRegister.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/Preprocessor.h>
#include <xloil/Log.h>
#include <xloil/LogWindowSink.h>
#include <spdlog/sinks/basic_file_sink.h>

namespace xloil
{
  XLO_FUNC_START(xloLog(
    const ExcelObj& showWindow
  ))
  {
    spdlog::default_logger()->flush();

    if (showWindow.get<bool>(false))
      openLogWindow();
    // TODO: better to add the log file name to the addin context?
    // TODO: this only returns the main log file path - each addin context could have one
    for (auto& sink : spdlog::default_logger()->sinks())
    {
      auto p = dynamic_cast<spdlog::sinks::basic_file_sink_mt*>(sink.get());
      if (p)
        return returnValue(p->filename());
    }
    return returnValue(CellError::NA);
  }
  XLO_FUNC_END(xloLog).threadsafe()
    .help(L"Flushes the log and returns the location of the log file")
    .arg(L"showWindow", L"Opens the log window");
}