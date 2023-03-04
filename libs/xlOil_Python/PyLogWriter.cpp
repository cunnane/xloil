#include "PyCore.h"
#include "PyAddin.h"
#include "PyHelpers.h"
#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <xlOil/Interface.h>
#include <xlOilHelpers/Settings.h>
#include <pybind11/stl.h>

using std::shared_ptr;
using std::wstring_view;
using std::vector;
using std::wstring;
using std::string;
namespace py = pybind11;

#if PY_VERSION_HEX < 0x030B0000
inline auto PyFrame_GetLasti(PyFrameObject* frame) { return frame->f_lasti; }
#endif
#if PY_VERSION_HEX < 0x03090000
inline auto PyFrame_GetCode(PyFrameObject* frame) { return frame->f_code; }
#endif

namespace xloil
{
  namespace Python
  {
    namespace
    {
      auto sourceFromFrame()
      {
        spdlog::source_loc source{ __FILE__, __LINE__, SPDLOG_FUNCTION };
        auto frame = PyEval_GetFrame();
        if (frame)
        {
          PyCodeObject* code = PyFrame_GetCode(frame); // Guaranteed never null
          source.line = PyCode_Addr2Line(code, PyFrame_GetLasti(frame));
          source.filename = PyUnicode_AsUTF8(code->co_filename);
          source.funcname = PyUnicode_AsUTF8(code->co_name);
        }
        return source;
      }

      /// <summary>
      /// Allows intial match like 'warn' for 'warning'
      /// </summary>
      /// <param name="target"></param>
      /// <returns></returns>
      spdlog::level::level_enum levelFromStr(const std::string& target)
      {
        using namespace spdlog::level;
        int iLevel = 0;
        for (const auto& level_str : SPDLOG_LEVEL_NAMES)
        {
          if (strncmp(target.c_str(), level_str, target.length()) == 0)
            return static_cast<level_enum>(iLevel);
          iLevel++;
        }
        return off;
      }

      spdlog::level::level_enum toSpdLogLevel(const py::object& level)
      {
        if (PyLong_Check(level.ptr()))
        {
          return spdlog::level::level_enum(
            std::min(PyLong_AsUnsignedLong(level.ptr()) / 10, 6ul));
        }
        return levelFromStr(toLower((string)py::str(level)));
      }

      // The numerical values of the python log levels align nicely with spdlog
      // so we can translate with a factor of 10.
      // https://docs.python.org/3/library/logging.html#levels

      class LogWriter
      {
      public:
        void writeToLog(const py::object& msg, const py::args& args, const py::kwargs& kwargs)
        {
          spdlog::level::level_enum level = spdlog::level::info;
          if (kwargs.contains("level"))
            level = toSpdLogLevel(kwargs["level"]);

          if (!spdlog::default_logger_raw()->should_log(level))
            return;

          if (kwargs.contains("file"))
          {
            auto file = to_string(kwargs["file"].ptr());
            auto func = to_string(kwargs["func"].ptr());
            auto line = py::cast<int>(kwargs["line"]);

            auto source = spdlog::source_loc(file.c_str(), line, func.c_str());
            writeToLogImpl(msg, args, level, source);
          }
          else
          {
            auto source = sourceFromFrame();
            writeToLogImpl(msg, args, level, source);
          }
        }

        void writeToLogHelper(
          const py::object& message, 
          const py::args& args, 
          spdlog::level::level_enum level)
        {
          if (!spdlog::default_logger_raw()->should_log(level))
            return;
          writeToLogImpl(message, args, level, sourceFromFrame());
        }

        void writeToLogImpl(
          const py::object& msg,
          const py::args& args, 
          spdlog::level::level_enum level,
          spdlog::source_loc source)
        {
          const string message = to_string(args.size() > 0
            ? PySteal<>(PyUnicode_Format(msg.ptr(), args.ptr()))
            : msg);
  
          py::gil_scoped_release releaseGil;
          spdlog::default_logger_raw()->log(
            source,
            level,
            message);
        }

        void flush()
        {
          py::gil_scoped_release releaseGil;
          spdlog::default_logger()->flush();
        }

        void trace(const py::object& msg, const py::args& args) { writeToLogHelper(msg, args, spdlog::level::trace); }
        void debug(const py::object& msg, const py::args& args) { writeToLogHelper(msg, args, spdlog::level::debug); }
        void info(const py::object& msg,  const py::args& args) { writeToLogHelper(msg, args, spdlog::level::info); }
        void warn(const py::object& msg,  const py::args& args) { writeToLogHelper(msg, args, spdlog::level::warn); }
        void error(const py::object& msg, const py::args& args) { writeToLogHelper(msg, args, spdlog::level::err); }

        auto getLogLevel() const
        {
          const char* levelNames[] SPDLOG_LEVEL_NAMES;
          auto level = spdlog::default_logger()->level();
          return levelNames[level];
        }

        unsigned getLogLevelInt() const
        {
          auto level = spdlog::default_logger()->level();
          return level * 10;
        }

        void setLogLevel(const py::object& level)
        {
          spdlog::default_logger()->set_level(toSpdLogLevel(level));
        }

        auto getFlushLevel() const
        {
          const char* levelNames[] SPDLOG_LEVEL_NAMES;
          auto level = spdlog::default_logger()->flush_level();
          return levelNames[level];
        }

        void setFlushLevel(const py::object& level)
        {
          spdlog::default_logger()->flush_on(toSpdLogLevel(level));
        }

        auto levels() const
        {
          // TODO: could be a static pyobject
          return vector<string> SPDLOG_LEVEL_NAMES;
        }

        auto logFilePath() const
        {
          return theCoreAddin()->context.logFilePath;
        }
      };

      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<LogWriter>(mod,
          "_LogWriter", R"(
            Writes a log message to xlOil's log.  The level parameter can be a level constant 
            from the `logging` module or one of the strings *error*, *warn*, *info*, *debug* or *trace*.

            Only messages with a level higher than the xlOil log level which is initially set
            to the value in the xlOil settings will be output to the log file. Trace output
            can only be seen with a debug build of xlOil.
          )")
         .def(py::init<>(), R"(
            Do not construct this class - a singleton instance is created by xlOil.
          )")
         .def("__call__",
            &LogWriter::writeToLog,
            R"(
              Writes a message to the log at the specifed keyword paramter `level`. The default 
              level is 'info'.  The message can contain format specifiers which are expanded
              using any additional positional arguments. This allows for lazy contruction of the 
              log string like python's own 'logging' module.
            )",
            py::arg("msg"))
          .def("flush", &LogWriter::flush,
            R"(
              Forces a log file 'flush', i.e write pending log messages to the log file.
              For performance reasons the file is not by default flushed for every message.
            )")
          .def("trace", &LogWriter::trace,
            "Writes a log message at the 'trace' level",
            py::arg("msg"))
          .def("debug", &LogWriter::debug,
            "Writes a log message at the 'debug' level",
            py::arg("msg"))
          .def("info", &LogWriter::info,
            "Writes a log message at the 'info' level",
            py::arg("msg"))
          .def("warn", &LogWriter::warn,
            "Writes a log message at the 'warn' level",
            py::arg("msg"))
          .def("error", &LogWriter::error,
            "Writes a log message at the 'error' level",
            py::arg("msg"))
          .def_property("level",
            &LogWriter::getLogLevel,
            &LogWriter::setLogLevel,
            R"(
              Returns or sets the current log level. The returned value will always be an 
              integer corresponding to levels in the `logging` module.  The level can be
              set to an integer or one of the strings *error*, *warn*, *info*, *debug* or *trace*.
            )")
          .def_property_readonly("level_int", &LogWriter::getLogLevelInt,
            R"(
              Returns the log level as an integer corresponding to levels in the `logging` module.
              Useful if you want to condition some output based on being above a certain log
              level.
            )")
          .def_property_readonly("levels",
            &LogWriter::levels,
            "A list of the available log levels")
          .def_property("flush_on",
            &LogWriter::getFlushLevel,
            &LogWriter::setFlushLevel,
            R"(
              Returns or sets the log level which will trigger a 'flush', i.e a writing pending
              log messages to the log file.
            )")
          .def_property_readonly("path",
            &LogWriter::logFilePath,
            "The full pathname of the log file");
      });
    }
  }
}