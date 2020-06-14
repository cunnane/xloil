#pragma once

namespace Excel { struct _Application; }

namespace xloil
{
  void reconnectCOM();

  Excel::_Application& excelApp();

  /// <summary>
  /// Returns true if the workbook is open. Like all COM functions it should 
  /// only be called on the main thread. Possible race condition if it is not
  /// </summary>
  bool checkWorkbookIsOpen(const wchar_t* wbName);
}