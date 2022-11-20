#include <xloil/WindowsSlim.h>
#include <xloil/ExcelCall.h>
#include <xloil/StringUtils.h>
#include <xloil/StaticRegister.h>
#include <xloil/FuncSpec.h>
#include <xloil/Events.h>
#include <xloil/ExcelArray.h>
#include <xloil-XLL/FuncRegistry.h>
#include <xloil/State.h>
#include <xloilHelpers/Environment.h>
#include <xloilHelpers/GuidUtils.h>
#include <list>
#include <string>
#include <filesystem>

using std::string;
using std::wstring;
using std::make_shared;
using std::shared_ptr;

namespace
{
  const GUID theExcelDnaNamespaceGuid = 
    { 0x306D016E, 0xCCE8, 0x4861, { 0x9D, 0xA1, 0x51, 0xA2, 0x7C, 0xBE, 0x34, 0x1A} };

  // Excel-DNA looks for a function `RegistrationInfo_<GUID>` where GUID is generated
  // by the algorithm in stableGuidFromString with the ExcelDNA namespace and the XLL 
  // path. The RegistrationInfo function should take a version number a return a 
  // specially formatted array describing registered functions and their help strings
	static wstring RegistrationInfoName(wstring xllPath)
	{
    GUID guid;
    xloil::stableGuidFromString(guid, theExcelDnaNamespaceGuid, xllPath);
    return wstring(L"RegistrationInfo_") + xloil::guidToWString(guid, false);
	}

  // Multiple Intellisense servers can exist in the same Excel session, but only one
  // will be active.  If a server is running, that server is described by the 
  // environment variable EXCELDNA_INTELLISENSE_ACTIVE_SERVER which takes the form
  // `<xll path>,<serverId>,<version>`.  We are only interested in the *serverId*.
	wstring findActiveIntelliServer()
	{
    auto active = xloil::getEnvironmentVar(L"EXCELDNA_INTELLISENSE_ACTIVE_SERVER");
    if (active.empty())
      return active;
    auto comma = active.find_first_of(L',');
    auto serverId = active.substr(comma + 1, active.find_last_of(L',') - comma - 1);
		return serverId;
	}

  // Calling `IntelliSenseServerControl_<GUID>("REFRESH")` where <GUID> is a serverId
  // triggers the server to look for new functions in the various sources it searches
  // including calling `RegistrationInfo_xxx` for known XLL addins.  The serverId is 
  // determined by findActiveIntelliServer.
  void triggerIntellisenseRefresh()
  {
    auto activeServer = findActiveIntelliServer();
    if (!activeServer.empty())
      xloil::tryCallExcel(msxll::xlUDF, L"IntelliSenseServerControl_" + activeServer, L"REFRESH");
  }
}

namespace xloil
{
  namespace
  {
    std::list<shared_ptr<const FuncInfo>> thePendingFuncInfos;
    static int theIntellisenseInfoVersion = 0;
    static wstring theIntellisenseRegisteredXll;
  }

  void publishIntellisenseInfo(const std::shared_ptr<const FuncInfo>& info)
  {
    if (theIntellisenseRegisteredXll.empty())
      return;
    thePendingFuncInfos.push_back(info);
    triggerIntellisenseRefresh();
  }

  void publishIntellisenseInfo(const std::vector<std::shared_ptr<const FuncInfo>>& infos)
  {
    if (theIntellisenseRegisteredXll.empty())
      return;
    thePendingFuncInfos.insert(thePendingFuncInfos.end(), infos.begin(), infos.end());
    triggerIntellisenseRefresh();
  }

  // The RegistrationInfo function should take a version number a return a specially 
  // formatted array describing registered functions and their help strings.  
  // The version number passed in is the version number returned by the previous call
  // (or -1 for the first call). We do not use this version number to keep track of
  // registrations, rather keeping a queue of pending info. However, the returned version
  // number must be larger that the supplied one or Excel-DNA will not process the 
  // returned array.
  // 
  // The returned array is sparse: the column positions correspond to argument numbers
  // for xlfRegister, but Excel-DNA only requires some of xlfRegister arguments to 
  // display the help.
  XLO_ENTRY_POINT(ExcelObj*) IntellisenseRegistrationInfo(const ExcelObj& /*version*/)
  {
    try
    {
      if (thePendingFuncInfos.empty())
        return returnValue(CellError::Num); // signals that no update is required

      // Do two passes through the function names because the array builder 
      // needs all the sizes upfront.  The strings are the XLL path, 
      // then each function name, category, help string, argument names
      // and argument help.
      size_t totalStrLen = theIntellisenseRegisteredXll.size();
      size_t maxNumArgs = 0;
      ExcelArray::row_t nFuncs = 0;
      for (auto& info : thePendingFuncInfos)
      {
        totalStrLen += info->name.size();
        totalStrLen += info->category.size();
        totalStrLen += info->help.size();
        for (auto x : info->args)
        {
          totalStrLen += x.name.size() + 1 + 2; // allow for comma and []
          totalStrLen += x.help.size();
        }
        maxNumArgs = std::max(maxNumArgs, info->args.size());
        ++nFuncs;
      }

      ExcelArrayBuilder block(1 + nFuncs, 10 + (ExcelArray::col_t)maxNumArgs, totalStrLen);
      block.fillNA();

      block(0, 0) = theIntellisenseRegisteredXll;
      block(0, 1) = theIntellisenseInfoVersion;

      ExcelObj::row_t i = 1;
      for (auto& info : thePendingFuncInfos)
      {
        wstring argNames;
        for (auto x : info->args)
          if (x.type & FuncArg::Optional)
            argNames.append(formatStr(L"[%s],", x.name.c_str()));
          else
            argNames.append(x.name).append(L",");
        if (info->numArgs() > 0)
          argNames.pop_back();  // Delete final comma

        // The row must start with Empty/Nil or Excel-DNA will skip it
        block(i, 0) = ExcelType::Nil; 
        block(i, 3) = info->name;
        block(i, 4) = argNames;
        block(i, 8) = info->category;
        block(i, 9) = info->help;
        
        // Write argument help strings from position 10 onwards
        for (size_t iArg = 0; iArg < info->numArgs(); ++iArg)
          block(i, 10 + iArg) = info->args[iArg].help;

        ++i;
      }
      
      thePendingFuncInfos.clear();
      ++theIntellisenseInfoVersion;

      return returnValue(block.toExcelObj());
    }
    catch (const std::exception& err)
    {
      return xloil::returnValue(err);
    }
  }

  void registerIntellisenseHook(const wchar_t* xllPath)
  {
    XLO_REGISTER_LATER(IntellisenseRegistrationInfo)
      .name(RegistrationInfoName(xllPath))
      .arg(L"version").hidden();
    theIntellisenseRegisteredXll = xllPath;
  }
}