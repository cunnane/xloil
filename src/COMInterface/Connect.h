namespace Excel { struct _Application; }

namespace xloil
{
  void reconnectCOM();
  Excel::_Application& excelApp();
}