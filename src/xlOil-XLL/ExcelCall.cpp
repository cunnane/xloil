#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelObj.h>
#include <cassert>
using namespace msxll;
using std::string;

namespace xloil
{
  const wchar_t* xlRetCodeToString(int ret, bool checkXllContext)
  {
    if (checkXllContext)
    {
      ExcelObj dummy;
      if (Excel12v(xlStack, &dummy, 0, nullptr) == xlretInvXlfn)
        return L"XLL function called outside XLL Context";
    }
    switch (ret)
    {
    case xlretSuccess:    return L"success";
    case xlretAbort:      return L"macro was stopped by the user ";
    case xlretInvXlfn:    return L"invalid function number, or calling function does not have permission to call the function or command";
    case xlretInvCount:   return L"invalid number of arguments";
    case xlretInvXloper:  return L"invalid XLOPER structure";
    case xlretStackOvfl:  return L"stack overflow";
    case xlretFailed:     return L"command failed";
    case xlretUncalced:   return L"attempt to read an uncalculated cell: this requires macro sheet permission";
    case xlretNotThreadSafe:  return L"not allowed during multi-threaded calc";
    case xlretInvAsynchronousContext: return L"invalid asynchronous function handle";
    case xlretNotClusterSafe:  return L"not supported on cluster";
    default:
      return L"unknown error";
    }
  }
  
  // TODO: currently unused, supposed to indicate functions which are safe to call outside XLL context
  // I'm not sure they are in fact safe!
  bool isSafeFunction(int funcNumber)
  {
    switch (funcNumber)
    {
    case xlFree:
    case xlStack:
    case xlSheetId:
    case xlSheetNm:
    case xlGetInst:
    case xlGetHwnd:
    case xlGetInstPtr:
    case xlAsyncReturn:
      return true;
    default:
      return false;
    }
  }

  XLOIL_EXPORT int callExcelRaw(
    int func, ExcelObj* result, size_t nArgs, const ExcelObj** args)
  {
    auto ret = Excel12v(func, result, (int)nArgs, (XLOIL_XLOPER**)args);
    if (result)
      result->fromExcel();
    return ret;
  }

  struct FuncNames 
  {
    static const char* names[xlfEncodeurl + 1];

    int lookup(const char* fname)
    {
      // Make the name lowercase and transform any dots to underscores
      auto name = string(fname);
      std::transform(name.begin(), name.end(), name.begin(), [](char c) 
      { 
        return c == '.' ? '_' : (char)std::tolower(c); 
      });

      auto i = std::lower_bound(sortOrder, sortOrder + _countof(sortOrder), name,
        [](int l, const string& r) { return strncmp(names[l], r.c_str(), r.length()) < 0; });
      return strcmp(names[*i], name.c_str()) == 0 ? *i : -1;
    }

    const char* name(unsigned funcNum) noexcept
    {
      if (funcNum >= _countof(names))
        return nullptr;
      return names[funcNum];
    }

    constexpr FuncNames()
    {
      names[xlfCount] = "count";
      names[1] = nullptr;
      names[xlfIsna] = "isna";
      names[xlfIserror] = "iserror";
      names[xlfSum] = "sum";
      names[xlfAverage] = "average";
      names[xlfMin] = "min";
      names[xlfMax] = "max";
      names[xlfRow] = "row";
      names[xlfColumn] = "column";
      names[xlfNa] = "na";
      names[xlfNpv] = "npv";
      names[xlfStdev] = "stdev";
      names[xlfDollar] = "dollar";
      names[xlfFixed] = "fixed";
      names[xlfSin] = "sin";
      names[xlfCos] = "cos";
      names[xlfTan] = "tan";
      names[xlfAtan] = "atan";
      names[xlfPi] = "pi";
      names[xlfSqrt] = "sqrt";
      names[xlfExp] = "exp";
      names[xlfLn] = "ln";
      names[xlfLog10] = "log10";
      names[xlfAbs] = "abs";
      names[xlfInt] = "int";
      names[xlfSign] = "sign";
      names[xlfRound] = "round";
      names[xlfLookup] = "lookup";
      names[xlfIndex] = "index";
      names[xlfRept] = "rept";
      names[xlfMid] = "mid";
      names[xlfLen] = "len";
      names[xlfValue] = "value";
      names[xlfTrue] = "true";
      names[xlfFalse] = "false";
      names[xlfAnd] = "and";
      names[xlfOr] = "or";
      names[xlfNot] = "not";
      names[xlfMod] = "mod";
      names[xlfDcount] = "dcount";
      names[xlfDsum] = "dsum";
      names[xlfDaverage] = "daverage";
      names[xlfDmin] = "dmin";
      names[xlfDmax] = "dmax";
      names[xlfDstdev] = "dstdev";
      names[xlfVar] = "var";
      names[xlfDvar] = "dvar";
      names[xlfText] = "text";
      names[xlfLinest] = "linest";
      names[xlfTrend] = "trend";
      names[xlfLogest] = "logest";
      names[xlfGrowth] = "growth";
      names[xlfGoto] = "goto";
      names[xlfHalt] = "halt";
      names[55] = nullptr;
      names[xlfPv] = "pv";
      names[xlfFv] = "fv";
      names[xlfNper] = "nper";
      names[xlfPmt] = "pmt";
      names[xlfRate] = "rate";
      names[xlfMirr] = "mirr";
      names[xlfIrr] = "irr";
      names[xlfRand] = "rand";
      names[xlfMatch] = "match";
      names[xlfDate] = "date";
      names[xlfTime] = "time";
      names[xlfDay] = "day";
      names[xlfMonth] = "month";
      names[xlfYear] = "year";
      names[xlfWeekday] = "weekday";
      names[xlfHour] = "hour";
      names[xlfMinute] = "minute";
      names[xlfSecond] = "second";
      names[xlfNow] = "now";
      names[xlfAreas] = "areas";
      names[xlfRows] = "rows";
      names[xlfColumns] = "columns";
      names[xlfOffset] = "offset";
      names[xlfAbsref] = "absref";
      names[xlfRelref] = "relref";
      names[xlfArgument] = "argument";
      names[xlfSearch] = "search";
      names[xlfTranspose] = "transpose";
      names[xlfError] = "error";
      names[xlfStep] = "step";
      names[xlfType] = "type";
      names[xlfEcho] = "echo";
      names[xlfSetName] = "setname";
      names[xlfCaller] = "caller";
      names[xlfDeref] = "deref";
      names[xlfWindows] = "windows";
      names[xlfSeries] = "series";
      names[xlfDocuments] = "documents";
      names[xlfActiveCell] = "activecell";
      names[xlfSelection] = "selection";
      names[xlfResult] = "result";
      names[xlfAtan2] = "atan2";
      names[xlfAsin] = "asin";
      names[xlfAcos] = "acos";
      names[xlfChoose] = "choose";
      names[xlfHlookup] = "hlookup";
      names[xlfVlookup] = "vlookup";
      names[xlfLinks] = "links";
      names[xlfInput] = "input";
      names[xlfIsref] = "isref";
      names[xlfGetFormula] = "getformula";
      names[xlfGetName] = "getname";
      names[xlfSetValue] = "setvalue";
      names[xlfLog] = "log";
      names[xlfExec] = "exec";
      names[xlfChar] = "char";
      names[xlfLower] = "lower";
      names[xlfUpper] = "upper";
      names[xlfProper] = "proper";
      names[xlfLeft] = "left";
      names[xlfRight] = "right";
      names[xlfExact] = "exact";
      names[xlfTrim] = "trim";
      names[xlfReplace] = "replace";
      names[xlfSubstitute] = "substitute";
      names[xlfCode] = "code";
      names[xlfNames] = "names";
      names[xlfDirectory] = "directory";
      names[xlfFind] = "find";
      names[xlfCell] = "cell";
      names[xlfIserr] = "iserr";
      names[xlfIstext] = "istext";
      names[xlfIsnumber] = "isnumber";
      names[xlfIsblank] = "isblank";
      names[xlfT] = "t";
      names[xlfN] = "n";
      names[xlfFopen] = "fopen";
      names[xlfFclose] = "fclose";
      names[xlfFsize] = "fsize";
      names[xlfFreadln] = "freadln";
      names[xlfFread] = "fread";
      names[xlfFwriteln] = "fwriteln";
      names[xlfFwrite] = "fwrite";
      names[xlfFpos] = "fpos";
      names[xlfDatevalue] = "datevalue";
      names[xlfTimevalue] = "timevalue";
      names[xlfSln] = "sln";
      names[xlfSyd] = "syd";
      names[xlfDdb] = "ddb";
      names[xlfGetDef] = "getdef";
      names[xlfReftext] = "reftext";
      names[xlfTextref] = "textref";
      names[xlfIndirect] = "indirect";
      names[xlfRegister] = "register";
      names[xlfCall] = "call";
      names[xlfAddBar] = "addbar";
      names[xlfAddMenu] = "addmenu";
      names[xlfAddCommand] = "addcommand";
      names[xlfEnableCommand] = "enablecommand";
      names[xlfCheckCommand] = "checkcommand";
      names[xlfRenameCommand] = "renamecommand";
      names[xlfShowBar] = "showbar";
      names[xlfDeleteMenu] = "deletemenu";
      names[xlfDeleteCommand] = "deletecommand";
      names[xlfGetChartItem] = "getchartitem";
      names[xlfDialogBox] = "dialogbox";
      names[xlfClean] = "clean";
      names[xlfMdeterm] = "mdeterm";
      names[xlfMinverse] = "minverse";
      names[xlfMmult] = "mmult";
      names[xlfFiles] = "files";
      names[xlfIpmt] = "ipmt";
      names[xlfPpmt] = "ppmt";
      names[xlfCounta] = "counta";
      names[xlfCancelKey] = "cancelkey";
      names[171] = nullptr;
      names[172] = nullptr;
      names[173] = nullptr;
      names[174] = nullptr;
      names[xlfInitiate] = "initiate";
      names[xlfRequest] = "request";
      names[xlfPoke] = "poke";
      names[xlfExecute] = "execute";
      names[xlfTerminate] = "terminate";
      names[xlfRestart] = "restart";
      names[xlfHelp] = "help";
      names[xlfGetBar] = "getbar";
      names[xlfProduct] = "product";
      names[xlfFact] = "fact";
      names[xlfGetCell] = "getcell";
      names[xlfGetWorkspace] = "getworkspace";
      names[xlfGetWindow] = "getwindow";
      names[xlfGetDocument] = "getdocument";
      names[xlfDproduct] = "dproduct";
      names[xlfIsnontext] = "isnontext";
      names[xlfGetNote] = "getnote";
      names[xlfNote] = "note";
      names[xlfStdevp] = "stdevp";
      names[xlfVarp] = "varp";
      names[xlfDstdevp] = "dstdevp";
      names[xlfDvarp] = "dvarp";
      names[xlfTrunc] = "trunc";
      names[xlfIslogical] = "islogical";
      names[xlfDcounta] = "dcounta";
      names[xlfDeleteBar] = "deletebar";
      names[xlfUnregister] = "unregister";
      names[202] = nullptr;
      names[203] = nullptr;
      names[xlfUsdollar] = "usdollar";
      names[xlfFindb] = "findb";
      names[xlfSearchb] = "searchb";
      names[xlfReplaceb] = "replaceb";
      names[xlfLeftb] = "leftb";
      names[xlfRightb] = "rightb";
      names[xlfMidb] = "midb";
      names[xlfLenb] = "lenb";
      names[xlfRoundup] = "roundup";
      names[xlfRounddown] = "rounddown";
      names[xlfAsc] = "asc";
      names[xlfDbcs] = "dbcs";
      names[xlfRank] = "rank";
      names[217] = nullptr;
      names[218] = nullptr;
      names[xlfAddress] = "address";
      names[xlfDays360] = "days360";
      names[xlfToday] = "today";
      names[xlfVdb] = "vdb";
      names[223] = nullptr;
      names[224] = nullptr;
      names[225] = nullptr;
      names[226] = nullptr;
      names[xlfMedian] = "median";
      names[xlfSumproduct] = "sumproduct";
      names[xlfSinh] = "sinh";
      names[xlfCosh] = "cosh";
      names[xlfTanh] = "tanh";
      names[xlfAsinh] = "asinh";
      names[xlfAcosh] = "acosh";
      names[xlfAtanh] = "atanh";
      names[xlfDget] = "dget";
      names[xlfCreateObject] = "createobject";
      names[xlfVolatile] = "volatile";
      names[xlfLastError] = "lasterror";
      names[xlfCustomUndo] = "customundo";
      names[xlfCustomRepeat] = "customrepeat";
      names[xlfFormulaConvert] = "formulaconvert";
      names[xlfGetLinkInfo] = "getlinkinfo";
      names[xlfTextBox] = "textbox";
      names[xlfInfo] = "info";
      names[xlfGroup] = "group";
      names[xlfGetObject] = "getobject";
      names[xlfDb] = "db";
      names[xlfPause] = "pause";
      names[249] = nullptr;
      names[250] = nullptr;
      names[xlfResume] = "resume";
      names[xlfFrequency] = "frequency";
      names[xlfAddToolbar] = "addtoolbar";
      names[xlfDeleteToolbar] = "deletetoolbar";
      names[xlUDF] = "udf";
      names[xlfResetToolbar] = "resettoolbar";
      names[xlfEvaluate] = "evaluate";
      names[xlfGetToolbar] = "gettoolbar";
      names[xlfGetTool] = "gettool";
      names[xlfSpellingCheck] = "spellingcheck";
      names[xlfErrorType] = "errortype";
      names[xlfAppTitle] = "apptitle";
      names[xlfWindowTitle] = "windowtitle";
      names[xlfSaveToolbar] = "savetoolbar";
      names[xlfEnableTool] = "enabletool";
      names[xlfPressTool] = "presstool";
      names[xlfRegisterId] = "registerid";
      names[xlfGetWorkbook] = "getworkbook";
      names[xlfAvedev] = "avedev";
      names[xlfBetadist] = "betadist";
      names[xlfGammaln] = "gammaln";
      names[xlfBetainv] = "betainv";
      names[xlfBinomdist] = "binomdist";
      names[xlfChidist] = "chidist";
      names[xlfChiinv] = "chiinv";
      names[xlfCombin] = "combin";
      names[xlfConfidence] = "confidence";
      names[xlfCritbinom] = "critbinom";
      names[xlfEven] = "even";
      names[xlfExpondist] = "expondist";
      names[xlfFdist] = "fdist";
      names[xlfFinv] = "finv";
      names[xlfFisher] = "fisher";
      names[xlfFisherinv] = "fisherinv";
      names[xlfFloor] = "floor";
      names[xlfGammadist] = "gammadist";
      names[xlfGammainv] = "gammainv";
      names[xlfCeiling] = "ceiling";
      names[xlfHypgeomdist] = "hypgeomdist";
      names[xlfLognormdist] = "lognormdist";
      names[xlfLoginv] = "loginv";
      names[xlfNegbinomdist] = "negbinomdist";
      names[xlfNormdist] = "normdist";
      names[xlfNormsdist] = "normsdist";
      names[xlfNorminv] = "norminv";
      names[xlfNormsinv] = "normsinv";
      names[xlfStandardize] = "standardize";
      names[xlfOdd] = "odd";
      names[xlfPermut] = "permut";
      names[xlfPoisson] = "poisson";
      names[xlfTdist] = "tdist";
      names[xlfWeibull] = "weibull";
      names[xlfSumxmy2] = "sumxmy2";
      names[xlfSumx2my2] = "sumx2my2";
      names[xlfSumx2py2] = "sumx2py2";
      names[xlfChitest] = "chitest";
      names[xlfCorrel] = "correl";
      names[xlfCovar] = "covar";
      names[xlfForecast] = "forecast";
      names[xlfFtest] = "ftest";
      names[xlfIntercept] = "intercept";
      names[xlfPearson] = "pearson";
      names[xlfRsq] = "rsq";
      names[xlfSteyx] = "steyx";
      names[xlfSlope] = "slope";
      names[xlfTtest] = "ttest";
      names[xlfProb] = "prob";
      names[xlfDevsq] = "devsq";
      names[xlfGeomean] = "geomean";
      names[xlfHarmean] = "harmean";
      names[xlfSumsq] = "sumsq";
      names[xlfKurt] = "kurt";
      names[xlfSkew] = "skew";
      names[xlfZtest] = "ztest";
      names[xlfLarge] = "large";
      names[xlfSmall] = "small";
      names[xlfQuartile] = "quartile";
      names[xlfPercentile] = "percentile";
      names[xlfPercentrank] = "percentrank";
      names[xlfMode] = "mode";
      names[xlfTrimmean] = "trimmean";
      names[xlfTinv] = "tinv";
      names[333] = nullptr;
      names[xlfMovieCommand] = "moviecommand";
      names[xlfGetMovie] = "getmovie";
      names[xlfConcatenate] = "concatenate";
      names[xlfPower] = "power";
      names[xlfPivotAddData] = "pivotadddata";
      names[xlfGetPivotTable] = "getpivottable";
      names[xlfGetPivotField] = "getpivotfield";
      names[xlfGetPivotItem] = "getpivotitem";
      names[xlfRadians] = "radians";
      names[xlfDegrees] = "degrees";
      names[xlfSubtotal] = "subtotal";
      names[xlfSumif] = "sumif";
      names[xlfCountif] = "countif";
      names[xlfCountblank] = "countblank";
      names[xlfScenarioGet] = "scenarioget";
      names[xlfOptionsListsGet] = "optionslistsget";
      names[xlfIspmt] = "ispmt";
      names[xlfDatedif] = "datedif";
      names[xlfDatestring] = "datestring";
      names[xlfNumberstring] = "numberstring";
      names[xlfRoman] = "roman";
      names[xlfOpenDialog] = "opendialog";
      names[xlfSaveDialog] = "savedialog";
      names[xlfViewGet] = "viewget";
      names[xlfGetpivotdata] = "getpivotdata";
      names[xlfHyperlink] = "hyperlink";
      names[xlfPhonetic] = "phonetic";
      names[xlfAveragea] = "averagea";
      names[xlfMaxa] = "maxa";
      names[xlfMina] = "mina";
      names[xlfStdevpa] = "stdevpa";
      names[xlfVarpa] = "varpa";
      names[xlfStdeva] = "stdeva";
      names[xlfVara] = "vara";
      names[xlfBahttext] = "bahttext";
      names[xlfThaidayofweek] = "thaidayofweek";
      names[xlfThaidigit] = "thaidigit";
      names[xlfThaimonthofyear] = "thaimonthofyear";
      names[xlfThainumsound] = "thainumsound";
      names[xlfThainumstring] = "thainumstring";
      names[xlfThaistringlength] = "thaistringlength";
      names[xlfIsthaidigit] = "isthaidigit";
      names[xlfRoundbahtdown] = "roundbahtdown";
      names[xlfRoundbahtup] = "roundbahtup";
      names[xlfThaiyear] = "thaiyear";
      names[xlfRtd] = "rtd";
      names[xlfCubevalue] = "cubevalue";
      names[xlfCubemember] = "cubemember";
      names[xlfCubememberproperty] = "cubememberproperty";
      names[xlfCuberankedmember] = "cuberankedmember";
      names[xlfHex2bin] = "hex2bin";
      names[xlfHex2dec] = "hex2dec";
      names[xlfHex2oct] = "hex2oct";
      names[xlfDec2bin] = "dec2bin";
      names[xlfDec2hex] = "dec2hex";
      names[xlfDec2oct] = "dec2oct";
      names[xlfOct2bin] = "oct2bin";
      names[xlfOct2hex] = "oct2hex";
      names[xlfOct2dec] = "oct2dec";
      names[xlfBin2dec] = "bin2dec";
      names[xlfBin2oct] = "bin2oct";
      names[xlfBin2hex] = "bin2hex";
      names[xlfImsub] = "imsub";
      names[xlfImdiv] = "imdiv";
      names[xlfImpower] = "impower";
      names[xlfImabs] = "imabs";
      names[xlfImsqrt] = "imsqrt";
      names[xlfImln] = "imln";
      names[xlfImlog2] = "imlog2";
      names[xlfImlog10] = "imlog10";
      names[xlfImsin] = "imsin";
      names[xlfImcos] = "imcos";
      names[xlfImexp] = "imexp";
      names[xlfImargument] = "imargument";
      names[xlfImconjugate] = "imconjugate";
      names[xlfImaginary] = "imaginary";
      names[xlfImreal] = "imreal";
      names[xlfComplex] = "complex";
      names[xlfImsum] = "imsum";
      names[xlfImproduct] = "improduct";
      names[xlfSeriessum] = "seriessum";
      names[xlfFactdouble] = "factdouble";
      names[xlfSqrtpi] = "sqrtpi";
      names[xlfQuotient] = "quotient";
      names[xlfDelta] = "delta";
      names[xlfGestep] = "gestep";
      names[xlfIseven] = "iseven";
      names[xlfIsodd] = "isodd";
      names[xlfMround] = "mround";
      names[xlfErf] = "erf";
      names[xlfErfc] = "erfc";
      names[xlfBesselj] = "besselj";
      names[xlfBesselk] = "besselk";
      names[xlfBessely] = "bessely";
      names[xlfBesseli] = "besseli";
      names[xlfXirr] = "xirr";
      names[xlfXnpv] = "xnpv";
      names[xlfPricemat] = "pricemat";
      names[xlfYieldmat] = "yieldmat";
      names[xlfIntrate] = "intrate";
      names[xlfReceived] = "received";
      names[xlfDisc] = "disc";
      names[xlfPricedisc] = "pricedisc";
      names[xlfYielddisc] = "yielddisc";
      names[xlfTbilleq] = "tbilleq";
      names[xlfTbillprice] = "tbillprice";
      names[xlfTbillyield] = "tbillyield";
      names[xlfPrice] = "price";
      names[xlfYield] = "yield";
      names[xlfDollarde] = "dollarde";
      names[xlfDollarfr] = "dollarfr";
      names[xlfNominal] = "nominal";
      names[xlfEffect] = "effect";
      names[xlfCumprinc] = "cumprinc";
      names[xlfCumipmt] = "cumipmt";
      names[xlfEdate] = "edate";
      names[xlfEomonth] = "eomonth";
      names[xlfYearfrac] = "yearfrac";
      names[xlfCoupdaybs] = "coupdaybs";
      names[xlfCoupdays] = "coupdays";
      names[xlfCoupdaysnc] = "coupdaysnc";
      names[xlfCoupncd] = "coupncd";
      names[xlfCoupnum] = "coupnum";
      names[xlfCouppcd] = "couppcd";
      names[xlfDuration] = "duration";
      names[xlfMduration] = "mduration";
      names[xlfOddlprice] = "oddlprice";
      names[xlfOddlyield] = "oddlyield";
      names[xlfOddfprice] = "oddfprice";
      names[xlfOddfyield] = "oddfyield";
      names[xlfRandbetween] = "randbetween";
      names[xlfWeeknum] = "weeknum";
      names[xlfAmordegrc] = "amordegrc";
      names[xlfAmorlinc] = "amorlinc";
      names[xlfConvert] = "convert";
      names[xlfAccrint] = "accrint";
      names[xlfAccrintm] = "accrintm";
      names[xlfWorkday] = "workday";
      names[xlfNetworkdays] = "networkdays";
      names[xlfGcd] = "gcd";
      names[xlfMultinomial] = "multinomial";
      names[xlfLcm] = "lcm";
      names[xlfFvschedule] = "fvschedule";
      names[xlfCubekpimember] = "cubekpimember";
      names[xlfCubeset] = "cubeset";
      names[xlfCubesetcount] = "cubesetcount";
      names[xlfIferror] = "iferror";
      names[xlfCountifs] = "countifs";
      names[xlfSumifs] = "sumifs";
      names[xlfAverageif] = "averageif";
      names[xlfAverageifs] = "averageifs";
      names[xlfAggregate] = "aggregate";
      names[xlfBinom_dist] = "binom_dist";
      names[xlfBinom_inv] = "binom_inv";
      names[xlfConfidence_norm] = "confidence_norm";
      names[xlfConfidence_t] = "confidence_t";
      names[xlfChisq_test] = "chisq_test";
      names[xlfF_test] = "f_test";
      names[xlfCovariance_p] = "covariance_p";
      names[xlfCovariance_s] = "covariance_s";
      names[xlfExpon_dist] = "expon_dist";
      names[xlfGamma_dist] = "gamma_dist";
      names[xlfGamma_inv] = "gamma_inv";
      names[xlfMode_mult] = "mode_mult";
      names[xlfMode_sngl] = "mode_sngl";
      names[xlfNorm_dist] = "norm_dist";
      names[xlfNorm_inv] = "norm_inv";
      names[xlfPercentile_exc] = "percentile_exc";
      names[xlfPercentile_inc] = "percentile_inc";
      names[xlfPercentrank_exc] = "percentrank_exc";
      names[xlfPercentrank_inc] = "percentrank_inc";
      names[xlfPoisson_dist] = "poisson_dist";
      names[xlfQuartile_exc] = "quartile_exc";
      names[xlfQuartile_inc] = "quartile_inc";
      names[xlfRank_avg] = "rank_avg";
      names[xlfRank_eq] = "rank_eq";
      names[xlfStdev_s] = "stdev_s";
      names[xlfStdev_p] = "stdev_p";
      names[xlfT_dist] = "t_dist";
      names[xlfT_dist_2t] = "t_dist_2t";
      names[xlfT_dist_rt] = "t_dist_rt";
      names[xlfT_inv] = "t_inv";
      names[xlfT_inv_2t] = "t_inv_2t";
      names[xlfVar_s] = "var_s";
      names[xlfVar_p] = "var_p";
      names[xlfWeibull_dist] = "weibull_dist";
      names[xlfNetworkdays_intl] = "networkdays_intl";
      names[xlfWorkday_intl] = "workday_intl";
      names[xlfEcma_ceiling] = "ecma_ceiling";
      names[xlfIso_ceiling] = "iso_ceiling";
      names[524] = nullptr;
      names[xlfBeta_dist] = "beta_dist";
      names[xlfBeta_inv] = "beta_inv";
      names[xlfChisq_dist] = "chisq_dist";
      names[xlfChisq_dist_rt] = "chisq_dist_rt";
      names[xlfChisq_inv] = "chisq_inv";
      names[xlfChisq_inv_rt] = "chisq_inv_rt";
      names[xlfF_dist] = "f_dist";
      names[xlfF_dist_rt] = "f_dist_rt";
      names[xlfF_inv] = "f_inv";
      names[xlfF_inv_rt] = "f_inv_rt";
      names[xlfHypgeom_dist] = "hypgeom_dist";
      names[xlfLognorm_dist] = "lognorm_dist";
      names[xlfLognorm_inv] = "lognorm_inv";
      names[xlfNegbinom_dist] = "negbinom_dist";
      names[xlfNorm_s_dist] = "norm_s_dist";
      names[xlfNorm_s_inv] = "norm_s_inv";
      names[xlfT_test] = "t_test";
      names[xlfZ_test] = "z_test";
      names[xlfErf_precise] = "erf_precise";
      names[xlfErfc_precise] = "erfc_precise";
      names[xlfGammaln_precise] = "gammaln_precise";
      names[xlfCeiling_precise] = "ceiling_precise";
      names[xlfFloor_precise] = "floor_precise";
      names[xlfAcot] = "acot";
      names[xlfAcoth] = "acoth";
      names[xlfCot] = "cot";
      names[xlfCoth] = "coth";
      names[xlfCsc] = "csc";
      names[xlfCsch] = "csch";
      names[xlfSec] = "sec";
      names[xlfSech] = "sech";
      names[xlfImtan] = "imtan";
      names[xlfImcot] = "imcot";
      names[xlfImcsc] = "imcsc";
      names[xlfImcsch] = "imcsch";
      names[xlfImsec] = "imsec";
      names[xlfImsech] = "imsech";
      names[xlfBitand] = "bitand";
      names[xlfBitor] = "bitor";
      names[xlfBitxor] = "bitxor";
      names[xlfBitlshift] = "bitlshift";
      names[xlfBitrshift] = "bitrshift";
      names[xlfPermutationa] = "permutationa";
      names[xlfCombina] = "combina";
      names[xlfXor] = "xor";
      names[xlfPduration] = "pduration";
      names[xlfBase] = "base";
      names[xlfDecimal] = "decimal";
      names[xlfDays] = "days";
      names[xlfBinom_dist_range] = "binom_dist_range";
      names[xlfGamma] = "gamma";
      names[xlfSkew_p] = "skew_p";
      names[xlfGauss] = "gauss";
      names[xlfPhi] = "phi";
      names[xlfRri] = "rri";
      names[xlfUnichar] = "unichar";
      names[xlfUnicode] = "unicode";
      names[xlfMunit] = "munit";
      names[xlfArabic] = "arabic";
      names[xlfIsoweeknum] = "isoweeknum";
      names[xlfNumbervalue] = "numbervalue";
      names[xlfSheet] = "sheet";
      names[xlfSheets] = "sheets";
      names[xlfFormulatext] = "formulatext";
      names[xlfIsformula] = "isformula";
      names[xlfIfna] = "ifna";
      names[xlfCeiling_math] = "ceiling_math";
      names[xlfFloor_math] = "floor_math";
      names[xlfImsinh] = "imsinh";
      names[xlfImcosh] = "imcosh";
      names[xlfFilterxml] = "filterxml";
      names[xlfWebservice] = "webservice";
      names[xlfEncodeurl] = "encodeurl";
    }

    static constexpr int sortOrder[] =
    {
      24,
      79,
      469,
      470,
      99,
      233,
      548,
      549,
      94,
      151,
      153,
      152,
      219,
      253,
      485,
      466,
      467,
      36,
      262,
      583,
      75,
      81,
      214,
      98,
      232,
      18,
      97,
      234,
      269,
      5,
      361,
      483,
      484,
      368,
      571,
      428,
      425,
      426,
      427,
      525,
      526,
      270,
      272,
      393,
      395,
      394,
      486,
      574,
      487,
      273,
      562,
      565,
      563,
      566,
      564,
      150,
      89,
      170,
      288,
      591,
      546,
      125,
      111,
      155,
      274,
      275,
      527,
      528,
      529,
      530,
      490,
      306,
      100,
      162,
      121,
      9,
      77,
      276,
      568,
      411,
      336,
      277,
      488,
      489,
      468,
      307,
      16,
      230,
      550,
      551,
      0,
      169,
      347,
      346,
      481,
      452,
      453,
      454,
      455,
      456,
      457,
      308,
      492,
      493,
      236,
      278,
      552,
      553,
      477,
      381,
      382,
      383,
      478,
      479,
      380,
      448,
      447,
      240,
      239,
      65,
      351,
      352,
      140,
      42,
      67,
      573,
      220,
      247,
      215,
      40,
      199,
      144,
      387,
      388,
      389,
      572,
      343,
      200,
      159,
      158,
      254,
      418,
      90,
      318,
      235,
      161,
      123,
      435,
      44,
      43,
      93,
      13,
      443,
      444,
      189,
      45,
      195,
      41,
      458,
      47,
      196,
      87,
      522,
      449,
      446,
      154,
      265,
      597,
      450,
      423,
      543,
      424,
      544,
      84,
      261,
      257,
      279,
      117,
      110,
      178,
      21,
      494,
      280,
      531,
      532,
      533,
      534,
      491,
      184,
      415,
      35,
      133,
      281,
      166,
      595,
      124,
      205,
      282,
      283,
      284,
      14,
      285,
      592,
      547,
      132,
      309,
      241,
      588,
      139,
      136,
      135,
      252,
      134,
      310,
      57,
      476,
      138,
      137,
      575,
      495,
      496,
      286,
      287,
      271,
      545,
      577,
      473,
      319,
      419,
      182,
      185,
      160,
      145,
      188,
      106,
      242,
      335,
      107,
      191,
      246,
      358,
      340,
      341,
      339,
      259,
      258,
      187,
      268,
      186,
      53,
      245,
      52,
      54,
      320,
      181,
      384,
      385,
      386,
      101,
      71,
      359,
      535,
      289,
      480,
      590,
      399,
      409,
      407,
      408,
      405,
      594,
      557,
      558,
      559,
      397,
      406,
      401,
      403,
      402,
      398,
      413,
      410,
      560,
      561,
      404,
      593,
      400,
      396,
      412,
      556,
      29,
      148,
      244,
      175,
      104,
      25,
      311,
      433,
      167,
      62,
      129,
      126,
      3,
      420,
      589,
      198,
      2,
      190,
      128,
      523,
      421,
      584,
      350,
      105,
      127,
      375,
      322,
      325,
      238,
      475,
      115,
      208,
      32,
      211,
      49,
      103,
      22,
      109,
      23,
      51,
      291,
      536,
      537,
      290,
      28,
      112,
      64,
      7,
      362,
      163,
      459,
      227,
      31,
      210,
      6,
      363,
      72,
      164,
      61,
      165,
      39,
      330,
      497,
      498,
      68,
      334,
      422,
      474,
      582,
      131,
      10,
      122,
      538,
      292,
      472,
      520,
      445,
      499,
      500,
      539,
      540,
      293,
      295,
      294,
      296,
      38,
      192,
      74,
      58,
      11,
      353,
      585,
      390,
      392,
      391,
      298,
      462,
      463,
      460,
      461,
      78,
      355,
      349,
      37,
      248,
      570,
      312,
      328,
      501,
      502,
      329,
      503,
      504,
      299,
      567,
      578,
      360,
      19,
      338,
      59,
      300,
      505,
      177,
      337,
      168,
      266,
      441,
      436,
      431,
      317,
      183,
      114,
      56,
      327,
      506,
      507,
      417,
      342,
      63,
      464,
      216,
      508,
      509,
      60,
      434,
      146,
      149,
      267,
      80,
      156,
      119,
      207,
      30,
      176,
      256,
      180,
      96,
      251,
      116,
      209,
      354,
      27,
      376,
      377,
      213,
      212,
      8,
      76,
      579,
      313,
      379,
      356,
      264,
      348,
      82,
      206,
      554,
      555,
      73,
      95,
      92,
      414,
      88,
      108,
      586,
      587,
      157,
      26,
      15,
      229,
      323,
      576,
      142,
      315,
      326,
      260,
      20,
      416,
      297,
      12,
      511,
      510,
      366,
      193,
      364,
      85,
      314,
      120,
      344,
      4,
      345,
      482,
      228,
      321,
      304,
      305,
      303,
      143,
      130,
      512,
      513,
      514,
      515,
      516,
      541,
      17,
      231,
      438,
      439,
      440,
      301,
      179,
      48,
      243,
      147,
      369,
      370,
      371,
      372,
      373,
      374,
      378,
      66,
      141,
      332,
      221,
      83,
      50,
      118,
      331,
      34,
      197,
      316,
      86,
      580,
      581,
      201,
      113,
      204,
      33,
      46,
      518,
      517,
      367,
      194,
      365,
      222,
      357,
      102,
      237,
      596,
      70,
      465,
      302,
      519,
      91,
      263,
      471,
      521,
      429,
      430,
      569,
      69,
      451,
      442,
      437,
      432,
      542,
      324
    };
  };
  const char* FuncNames::names[xlfEncodeurl + 1];

  int excelFuncNumber(const char* name)
  {
    return FuncNames().lookup(name);
  }
  const char* excelFuncName(const unsigned number) noexcept
  {
    return FuncNames().name(number);
  }

}
