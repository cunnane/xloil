#pragma once
#include <xloil/Throw.h>

namespace Excel { struct _Application; }

namespace xloil 
{
  namespace COM
  {
    // We don't inherit from our own Exception class as that writes to the log in its ctor
    class ComConnectException : public std::runtime_error
    {
    public:
      ComConnectException(const char* message)
        : std::runtime_error(message)
      {}
    };

    void connectCom();
    void disconnectCom();

    bool isComApiAvailable() noexcept;

    Excel::_Application& excelApp();

    /// <summary>
    /// Returns true if the workbook is open. Like all COM functions it should 
    /// only be called on the main thread. Possible race condition if it is not
    /// </summary>
    bool checkWorkbookIsOpen(const wchar_t* wbName);
  }
}