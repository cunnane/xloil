#pragma once

/* EDITED TO REMOVE DEPENDENCY ON WINDOWS.H */

/*
**  Microsoft Excel Developer's Toolkit
**  Version 15.0
**
**  File:           INCLUDE\XLCALL.H
**  Description:    Header file for for Excel callbacks
**  Platform:       Microsoft Windows
**
**  DEPENDENCY: none
**
**  This file defines the constants and
**  data types which are used in the
**  Microsoft Excel C API.
**
*/

/*
** XL 12 Basic Datatypes
**/

namespace msxll
{
  typedef unsigned char BYTE;
  typedef unsigned short WORD;
  typedef unsigned long DWORD;

  typedef __int32 BOOL;			/* Boolean */
  typedef wchar_t XCHAR;			/* Wide Character */
  typedef __int32 RW;			/* XL 12 Row */
  typedef __int32 COL;	 	      	/* XL 12 Column */
  typedef DWORD* IDSHEET;		/* XL12 Sheet ID */


  /*
  ** XLREF structure
  **
  ** Describes a single rectangular reference.
  */

  typedef struct xlref
  {
    WORD rwFirst;
    WORD rwLast;
    BYTE colFirst;
    BYTE colLast;
  } XLREF, *LPXLREF;


  /*
  ** XLMREF structure
  **
  ** Describes multiple rectangular references.
  ** This is a variable size structure, default
  ** size is 1 reference.
  */

  typedef struct xlmref
  {
    WORD count;
    XLREF reftbl[1];					/* actually reftbl[count] */
  } XLMREF, *LPXLMREF;


  /*
  ** XLREF12 structure
  **
  ** Describes a single XL 12 rectangular reference.
  */

  typedef struct xlref12
  {
    RW rwFirst;
    RW rwLast;
    COL colFirst;
    COL colLast;
  } XLREF12, *LPXLREF12;


  /*
  ** XLMREF12 structure
  **
  ** Describes multiple rectangular XL 12 references.
  ** This is a variable size structure, default
  ** size is 1 reference.
  */

  typedef struct xlmref12
  {
    WORD count;
    XLREF12 reftbl[1];					/* actually reftbl[count] */
  } XLMREF12, *LPXLMREF12;


  /*
  ** FP structure
  **
  ** Describes FP structure.
  */

  typedef struct _FP
  {
    unsigned short int rows;
    unsigned short int columns;
    double array[1];        /* Actually, array[rows][columns] */
  } FP;

  /*
  ** FP12 structure
  **
  ** Describes FP structure capable of handling the big grid.
  */

  typedef struct _FP12
  {
    __int32 rows;
    __int32 columns;
    double array[1];        /* Actually, array[rows][columns] */
  } FP12;


  /*
  ** XLOPER structure
  **
  ** Excel's fundamental data type: can hold data
  ** of any type. Use "R" as the argument type in the
  ** REGISTER function.
  **/

  typedef struct xloper
  {
    union
    {
      double num;					/* xltypeNum */
      char* str;					/* xltypeStr */
#ifdef __cplusplus
      WORD xbool;					/* xltypeBool */
#else	
      WORD bool;					/* xltypeBool */
#endif	
      WORD err;					/* xltypeErr */
      short int w;					/* xltypeInt */
      struct
      {
        WORD count;				/* always = 1 */
        XLREF ref;
      } sref;						/* xltypeSRef */
      struct
      {
        XLMREF *lpmref;
        IDSHEET idSheet;
      } mref;						/* xltypeRef */
      struct
      {
        struct xloper *lparray;
        WORD rows;
        WORD columns;
      } array;					/* xltypeMulti */
      struct
      {
        union
        {
          short int level;		/* xlflowRestart */
          short int tbctrl;		/* xlflowPause */
          IDSHEET idSheet;		/* xlflowGoto */
        } valflow;
        WORD rw;				/* xlflowGoto */
        BYTE col;				/* xlflowGoto */
        BYTE xlflow;
      } flow;						/* xltypeFlow */
      struct
      {
        union
        {
          BYTE *lpbData;			/* data passed to XL */
          void* hdata;			/* data returned from XL */
        } h;
        long cbData;
      } bigdata;					/* xltypeBigData */
    } val;
    WORD xltype;
  } XLOPER, *LPXLOPER;

  /*
  ** XLOPER12 structure
  **
  ** Excel 12's fundamental data type: can hold data
  ** of any type. Use "U" as the argument type in the
  ** REGISTER function.
  **/

  typedef struct xloper12
  {
    union
    {
      double num;				       	/* xltypeNum */
      XCHAR *str;				       	/* xltypeStr */
      BOOL xbool;				       	/* xltypeBool */
      int err;				       	/* xltypeErr */
      int w;
      struct
      {
        WORD count;			       	/* always = 1 */
        XLREF12 ref;
      } sref;						/* xltypeSRef */
      struct
      {
        XLMREF12 *lpmref;
        IDSHEET idSheet;
      } mref;						/* xltypeRef */
      struct
      {
        struct xloper12 *lparray;
        RW rows;
        COL columns;
      } array;					/* xltypeMulti */
      struct
      {
        union
        {
          int level;			/* xlflowRestart */
          int tbctrl;			/* xlflowPause */
          IDSHEET idSheet;		/* xlflowGoto */
        } valflow;
        RW rw;				       	/* xlflowGoto */
        COL col;			       	/* xlflowGoto */
        BYTE xlflow;
      } flow;						/* xltypeFlow */
      struct
      {
        union
        {
          BYTE *lpbData;			/* data passed to XL */
          void* hdata;			/* data returned from XL */
        } h;
        long cbData;
      } bigdata;					/* xltypeBigData */
    } val;
    DWORD xltype;
  } XLOPER12, *LPXLOPER12;

  /*
  ** XLOPER and XLOPER12 data types
  **
  ** Used for xltype field of XLOPER and XLOPER12 structures
  */

  constexpr unsigned xltypeNum = 0x0001;
  constexpr unsigned xltypeStr = 0x0002;
  constexpr unsigned xltypeBool = 0x0004;
  constexpr unsigned xltypeRef = 0x0008;
  constexpr unsigned xltypeErr = 0x0010;
  constexpr unsigned xltypeFlow = 0x0020;
  constexpr unsigned xltypeMulti = 0x0040;
  constexpr unsigned xltypeMissing = 0x0080;
  constexpr unsigned xltypeNil = 0x0100;
  constexpr unsigned xltypeSRef = 0x0400;
  constexpr unsigned xltypeInt = 0x0800;

  constexpr unsigned xlbitXLFree = 0x1000;
  constexpr unsigned xlbitDLLFree = 0x4000;

  constexpr unsigned xltypeBigData = (xltypeStr | xltypeInt);


  /*
  ** Error codes
  **
  ** Used for val.err field of XLOPER and XLOPER12 structures
  ** when constructing error XLOPERs and XLOPER12s
  */

  constexpr unsigned xlerrNull = 0;
  constexpr unsigned xlerrDiv0 = 7;
  constexpr unsigned xlerrValue = 15;
  constexpr unsigned xlerrRef = 23;
  constexpr unsigned xlerrName = 29;
  constexpr unsigned xlerrNum = 36;
  constexpr unsigned xlerrNA = 42;
  constexpr unsigned xlerrGettingData = 43;

  /*
  ** Return codes
  **
  ** These values can be returned from Excel4(), Excel4v(), Excel12() or Excel12v().
  */

  constexpr unsigned xlretSuccess                = 0;    /* success */ 
  constexpr unsigned xlretAbort                  = 1;    /* macro halted */
  constexpr unsigned xlretInvXlfn                = 2;    /* invalid function number */ 
  constexpr unsigned xlretInvCount               = 4;    /* invalid number of arguments */ 
  constexpr unsigned xlretInvXloper              = 8;    /* invalid OPER structure */  
  constexpr unsigned xlretStackOvfl              = 16;   /* stack overflow */  
  constexpr unsigned xlretFailed                 = 32;   /* command failed */  
  constexpr unsigned xlretUncalced               = 64;   /* uncalced cell */
  constexpr unsigned xlretNotThreadSafe          = 128;  /* not allowed during multi-threaded calc */
  constexpr unsigned xlretInvAsynchronousContext = 256;  /* invalid asynchronous function handle */
  constexpr unsigned xlretNotClusterSafe         = 512;  /* not supported on cluster */


  /*
  ** XLL events
  **
  ** Passed in to an xlEventRegister call to register a corresponding event.
  */

#define xleventCalculationEnded      1    /* Fires at the end of calculation */ 
#define xleventCalculationCanceled   2    /* Fires when calculation is interrupted */


  /*
  ** Function prototypes
  */

#ifdef __cplusplus
  extern "C" {
#endif

    int _cdecl Excel4(int xlfn, LPXLOPER operRes, int count, ...);
    /* followed by count LPXLOPERs */

    int __stdcall Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[]);

    int __stdcall XLCallVer(void);

    long __stdcall LPenHelper(int wCode, void* lpv);

    int _cdecl Excel12(int xlfn, LPXLOPER12 operRes, int count, ...);
    /* followed by count LPXLOPER12s */

    int __stdcall Excel12v(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12 opers[]);

#ifdef __cplusplus
  }
#endif


  /*
  ** Cluster Connector Async Callback
  */

  typedef int(__stdcall *PXL_HPC_ASYNC_CALLBACK)(LPXLOPER12 operAsyncHandle, LPXLOPER12 operReturn);


  /*
  ** Cluster connector entry point return codes
  */

#define xlHpcRetSuccess            0
#define xlHpcRetSessionIdInvalid  -1
#define xlHpcRetCallFailed        -2


  /*
  ** Function number bits
  */

  constexpr unsigned xlCommand = 0x8000;
  constexpr unsigned xlSpecial = 0x4000;
  constexpr unsigned xlIntl = 0x2000;
  constexpr unsigned xlPrompt = 0x1000;


  /*
  ** Auxiliary function numbers
  **
  ** These functions are available only from the C API,
  ** not from the Excel macro language.
  */

  constexpr unsigned xlFree = (0 | xlSpecial);
  constexpr unsigned xlStack = (1 | xlSpecial);
  constexpr unsigned xlCoerce = (2 | xlSpecial);
  constexpr unsigned xlSet = (3 | xlSpecial);
  constexpr unsigned xlSheetId = (4 | xlSpecial);
  constexpr unsigned xlSheetNm = (5 | xlSpecial);
  constexpr unsigned xlAbort = (6 | xlSpecial);
  constexpr unsigned xlGetInst = (7 | xlSpecial); /* Returns application's hinstance as an integer value, supported on 32-bit platform only */
  constexpr unsigned xlGetHwnd = (8 | xlSpecial);
  constexpr unsigned xlGetName = (9 | xlSpecial);
  constexpr unsigned xlEnableXLMsgs = (10 | xlSpecial);
  constexpr unsigned xlDisableXLMsgs = (11 | xlSpecial);
  constexpr unsigned xlDefineBinaryName = (12 | xlSpecial);
  constexpr unsigned xlGetBinaryName = (13 | xlSpecial);
  /* GetFooInfo are valid only for calls to LPenHelper */
 constexpr unsigned xlGetFmlaInfo	= (14 | xlSpecial);
 constexpr unsigned xlGetMouseInfo = (15 | xlSpecial);
 constexpr unsigned xlAsyncReturn	= (16 | xlSpecial);	/*Set return value from an asynchronous function call*/
 constexpr unsigned xlEventRegister = (17 | xlSpecial);	/*Register an XLL event*/
 constexpr unsigned xlRunningOnCluster = (18 | xlSpecial);	/*Returns true if running on Compute Cluster*/
 constexpr unsigned xlGetInstPtr = (19 | xlSpecial);	/* Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms */

/* edit modes */
#define xlModeReady	0	// not in edit mode
#define xlModeEnter	1	// enter mode
#define xlModeEdit	2	// edit mode
#define xlModePoint	4	// point mode

/* document(page) types */
#define dtNil 0x7f	// window is not a sheet, macro, chart or basic
// OR window is not the selected window at idle state
#define dtSheet 0	// sheet
#define dtProc  1	// XLM macro
#define dtChart 2	// Chart
#define dtBasic 6	// VBA 

/* hit test codes */
#define htNone		0x00	// none of below
#define htClient	0x01	// internal for "in the client are", should never see
#define htVSplit	0x02	// vertical split area with split panes
#define htHSplit	0x03	// horizontal split area
#define htColWidth	0x04	// column width adjuster area
#define htRwHeight	0x05	// row height adjuster area
#define htRwColHdr	0x06	// the intersection of row and column headers
#define htObject	0x07	// the body of an object
// the following are for size handles of draw objects
#define htTopLeft	0x08
#define htBotLeft	0x09
#define htLeft		0x0A
#define htTopRight	0x0B
#define htBotRight	0x0C
#define htRight		0x0D
#define htTop		0x0E
#define htBot		0x0F
// end size handles
#define htRwGut		0x10	// row area of outline gutter
#define htColGut	0x11	// column area of outline gutter
#define htTextBox	0x12	// body of a text box (where we shouw I-Beam cursor)
#define htRwLevels	0x13	// row levels buttons of outline gutter
#define htColLevels	0x14	// column levels buttons of outline gutter
#define htDman		0x15	// the drag/drop handle of the selection
#define htDmanFill	0x16	// the auto-fill handle of the selection
#define htXSplit	0x17	// the intersection of the horz & vert pane splits
#define htVertex	0x18	// a vertex of a polygon draw object
#define htAddVtx	0x19	// htVertex in add a vertex mode
#define htDelVtx	0x1A	// htVertex in delete a vertex mode
#define htRwHdr		0x1B	// row header
#define htColHdr	0x1C	// column header
#define htRwShow	0x1D	// Like htRowHeight except means grow a hidden column
#define htColShow	0x1E	// column version of htRwShow
#define htSizing	0x1F	// Internal use only
#define htSxpivot	0x20	// a drag/drop tile in a pivot table
#define htTabs		0x21	// the sheet paging tabs
#define htEdit		0x22	// Internal use only

  typedef struct _fmlainfo
  {
    int wPointMode;	// current edit mode.  0 => rest of struct undefined
    int cch;	// count of characters in formula
    char *lpch;	// pointer to formula characters.  READ ONLY!!!
    int ichFirst;	// char offset to start of selection
    int ichLast;	// char offset to end of selection (may be > cch)
    int ichCaret;	// char offset to blinking caret
  } FMLAINFO;


  /*
  ** User defined function
  **
  ** First argument should be a function reference.
  */


  constexpr unsigned xlUDF = 255;


  // Excel function numbers

constexpr unsigned xlfCount = 0;
constexpr unsigned xlfIsna = 2;
constexpr unsigned xlfIserror = 3;
constexpr unsigned xlfSum = 4;
constexpr unsigned xlfAverage = 5;
constexpr unsigned xlfMin = 6;
constexpr unsigned xlfMax = 7;
constexpr unsigned xlfRow = 8;
constexpr unsigned xlfColumn = 9;
constexpr unsigned xlfNa = 10;
constexpr unsigned xlfNpv = 11;
constexpr unsigned xlfStdev = 12;
constexpr unsigned xlfDollar = 13;
constexpr unsigned xlfFixed = 14;
constexpr unsigned xlfSin = 15;
constexpr unsigned xlfCos = 16;
constexpr unsigned xlfTan = 17;
constexpr unsigned xlfAtan = 18;
constexpr unsigned xlfPi = 19;
constexpr unsigned xlfSqrt = 20;
constexpr unsigned xlfExp = 21;
constexpr unsigned xlfLn = 22;
constexpr unsigned xlfLog10 = 23;
constexpr unsigned xlfAbs = 24;
constexpr unsigned xlfInt = 25;
constexpr unsigned xlfSign = 26;
constexpr unsigned xlfRound = 27;
constexpr unsigned xlfLookup = 28;
constexpr unsigned xlfIndex = 29;
constexpr unsigned xlfRept = 30;
constexpr unsigned xlfMid = 31;
constexpr unsigned xlfLen = 32;
constexpr unsigned xlfValue = 33;
constexpr unsigned xlfTrue = 34;
constexpr unsigned xlfFalse = 35;
constexpr unsigned xlfAnd = 36;
constexpr unsigned xlfOr = 37;
constexpr unsigned xlfNot = 38;
constexpr unsigned xlfMod = 39;
constexpr unsigned xlfDcount = 40;
constexpr unsigned xlfDsum = 41;
constexpr unsigned xlfDaverage = 42;
constexpr unsigned xlfDmin = 43;
constexpr unsigned xlfDmax = 44;
constexpr unsigned xlfDstdev = 45;
constexpr unsigned xlfVar = 46;
constexpr unsigned xlfDvar = 47;
constexpr unsigned xlfText = 48;
constexpr unsigned xlfLinest = 49;
constexpr unsigned xlfTrend = 50;
constexpr unsigned xlfLogest = 51;
constexpr unsigned xlfGrowth = 52;
constexpr unsigned xlfGoto = 53;
constexpr unsigned xlfHalt = 54;
constexpr unsigned xlfPv = 56;
constexpr unsigned xlfFv = 57;
constexpr unsigned xlfNper = 58;
constexpr unsigned xlfPmt = 59;
constexpr unsigned xlfRate = 60;
constexpr unsigned xlfMirr = 61;
constexpr unsigned xlfIrr = 62;
constexpr unsigned xlfRand = 63;
constexpr unsigned xlfMatch = 64;
constexpr unsigned xlfDate = 65;
constexpr unsigned xlfTime = 66;
constexpr unsigned xlfDay = 67;
constexpr unsigned xlfMonth = 68;
constexpr unsigned xlfYear = 69;
constexpr unsigned xlfWeekday = 70;
constexpr unsigned xlfHour = 71;
constexpr unsigned xlfMinute = 72;
constexpr unsigned xlfSecond = 73;
constexpr unsigned xlfNow = 74;
constexpr unsigned xlfAreas = 75;
constexpr unsigned xlfRows = 76;
constexpr unsigned xlfColumns = 77;
constexpr unsigned xlfOffset = 78;
constexpr unsigned xlfAbsref = 79;
constexpr unsigned xlfRelref = 80;
constexpr unsigned xlfArgument = 81;
constexpr unsigned xlfSearch = 82;
constexpr unsigned xlfTranspose = 83;
constexpr unsigned xlfError = 84;
constexpr unsigned xlfStep = 85;
constexpr unsigned xlfType = 86;
constexpr unsigned xlfEcho = 87;
constexpr unsigned xlfSetName = 88;
constexpr unsigned xlfCaller = 89;
constexpr unsigned xlfDeref = 90;
constexpr unsigned xlfWindows = 91;
constexpr unsigned xlfSeries = 92;
constexpr unsigned xlfDocuments = 93;
constexpr unsigned xlfActiveCell = 94;
constexpr unsigned xlfSelection = 95;
constexpr unsigned xlfResult = 96;
constexpr unsigned xlfAtan2 = 97;
constexpr unsigned xlfAsin = 98;
constexpr unsigned xlfAcos = 99;
constexpr unsigned xlfChoose = 100;
constexpr unsigned xlfHlookup = 101;
constexpr unsigned xlfVlookup = 102;
constexpr unsigned xlfLinks = 103;
constexpr unsigned xlfInput = 104;
constexpr unsigned xlfIsref = 105;
constexpr unsigned xlfGetFormula = 106;
constexpr unsigned xlfGetName = 107;
constexpr unsigned xlfSetValue = 108;
constexpr unsigned xlfLog = 109;
constexpr unsigned xlfExec = 110;
constexpr unsigned xlfChar = 111;
constexpr unsigned xlfLower = 112;
constexpr unsigned xlfUpper = 113;
constexpr unsigned xlfProper = 114;
constexpr unsigned xlfLeft = 115;
constexpr unsigned xlfRight = 116;
constexpr unsigned xlfExact = 117;
constexpr unsigned xlfTrim = 118;
constexpr unsigned xlfReplace = 119;
constexpr unsigned xlfSubstitute = 120;
constexpr unsigned xlfCode = 121;
constexpr unsigned xlfNames = 122;
constexpr unsigned xlfDirectory = 123;
constexpr unsigned xlfFind = 124;
constexpr unsigned xlfCell = 125;
constexpr unsigned xlfIserr = 126;
constexpr unsigned xlfIstext = 127;
constexpr unsigned xlfIsnumber = 128;
constexpr unsigned xlfIsblank = 129;
constexpr unsigned xlfT = 130;
constexpr unsigned xlfN = 131;
constexpr unsigned xlfFopen = 132;
constexpr unsigned xlfFclose = 133;
constexpr unsigned xlfFsize = 134;
constexpr unsigned xlfFreadln = 135;
constexpr unsigned xlfFread = 136;
constexpr unsigned xlfFwriteln = 137;
constexpr unsigned xlfFwrite = 138;
constexpr unsigned xlfFpos = 139;
constexpr unsigned xlfDatevalue = 140;
constexpr unsigned xlfTimevalue = 141;
constexpr unsigned xlfSln = 142;
constexpr unsigned xlfSyd = 143;
constexpr unsigned xlfDdb = 144;
constexpr unsigned xlfGetDef = 145;
constexpr unsigned xlfReftext = 146;
constexpr unsigned xlfTextref = 147;
constexpr unsigned xlfIndirect = 148;
constexpr unsigned xlfRegister = 149;
constexpr unsigned xlfCall = 150;
constexpr unsigned xlfAddBar = 151;
constexpr unsigned xlfAddMenu = 152;
constexpr unsigned xlfAddCommand = 153;
constexpr unsigned xlfEnableCommand = 154;
constexpr unsigned xlfCheckCommand = 155;
constexpr unsigned xlfRenameCommand = 156;
constexpr unsigned xlfShowBar = 157;
constexpr unsigned xlfDeleteMenu = 158;
constexpr unsigned xlfDeleteCommand = 159;
constexpr unsigned xlfGetChartItem = 160;
constexpr unsigned xlfDialogBox = 161;
constexpr unsigned xlfClean = 162;
constexpr unsigned xlfMdeterm = 163;
constexpr unsigned xlfMinverse = 164;
constexpr unsigned xlfMmult = 165;
constexpr unsigned xlfFiles = 166;
constexpr unsigned xlfIpmt = 167;
constexpr unsigned xlfPpmt = 168;
constexpr unsigned xlfCounta = 169;
constexpr unsigned xlfCancelKey = 170;
constexpr unsigned xlfInitiate = 175;
constexpr unsigned xlfRequest = 176;
constexpr unsigned xlfPoke = 177;
constexpr unsigned xlfExecute = 178;
constexpr unsigned xlfTerminate = 179;
constexpr unsigned xlfRestart = 180;
constexpr unsigned xlfHelp = 181;
constexpr unsigned xlfGetBar = 182;
constexpr unsigned xlfProduct = 183;
constexpr unsigned xlfFact = 184;
constexpr unsigned xlfGetCell = 185;
constexpr unsigned xlfGetWorkspace = 186;
constexpr unsigned xlfGetWindow = 187;
constexpr unsigned xlfGetDocument = 188;
constexpr unsigned xlfDproduct = 189;
constexpr unsigned xlfIsnontext = 190;
constexpr unsigned xlfGetNote = 191;
constexpr unsigned xlfNote = 192;
constexpr unsigned xlfStdevp = 193;
constexpr unsigned xlfVarp = 194;
constexpr unsigned xlfDstdevp = 195;
constexpr unsigned xlfDvarp = 196;
constexpr unsigned xlfTrunc = 197;
constexpr unsigned xlfIslogical = 198;
constexpr unsigned xlfDcounta = 199;
constexpr unsigned xlfDeleteBar = 200;
constexpr unsigned xlfUnregister = 201;
constexpr unsigned xlfUsdollar = 204;
constexpr unsigned xlfFindb = 205;
constexpr unsigned xlfSearchb = 206;
constexpr unsigned xlfReplaceb = 207;
constexpr unsigned xlfLeftb = 208;
constexpr unsigned xlfRightb = 209;
constexpr unsigned xlfMidb = 210;
constexpr unsigned xlfLenb = 211;
constexpr unsigned xlfRoundup = 212;
constexpr unsigned xlfRounddown = 213;
constexpr unsigned xlfAsc = 214;
constexpr unsigned xlfDbcs = 215;
constexpr unsigned xlfRank = 216;
constexpr unsigned xlfAddress = 219;
constexpr unsigned xlfDays360 = 220;
constexpr unsigned xlfToday = 221;
constexpr unsigned xlfVdb = 222;
constexpr unsigned xlfMedian = 227;
constexpr unsigned xlfSumproduct = 228;
constexpr unsigned xlfSinh = 229;
constexpr unsigned xlfCosh = 230;
constexpr unsigned xlfTanh = 231;
constexpr unsigned xlfAsinh = 232;
constexpr unsigned xlfAcosh = 233;
constexpr unsigned xlfAtanh = 234;
constexpr unsigned xlfDget = 235;
constexpr unsigned xlfCreateObject = 236;
constexpr unsigned xlfVolatile = 237;
constexpr unsigned xlfLastError = 238;
constexpr unsigned xlfCustomUndo = 239;
constexpr unsigned xlfCustomRepeat = 240;
constexpr unsigned xlfFormulaConvert = 241;
constexpr unsigned xlfGetLinkInfo = 242;
constexpr unsigned xlfTextBox = 243;
constexpr unsigned xlfInfo = 244;
constexpr unsigned xlfGroup = 245;
constexpr unsigned xlfGetObject = 246;
constexpr unsigned xlfDb = 247;
constexpr unsigned xlfPause = 248;
constexpr unsigned xlfResume = 251;
constexpr unsigned xlfFrequency = 252;
constexpr unsigned xlfAddToolbar = 253;
constexpr unsigned xlfDeleteToolbar = 254;
constexpr unsigned xlfResetToolbar = 256;
constexpr unsigned xlfEvaluate = 257;
constexpr unsigned xlfGetToolbar = 258;
constexpr unsigned xlfGetTool = 259;
constexpr unsigned xlfSpellingCheck = 260;
constexpr unsigned xlfErrorType = 261;
constexpr unsigned xlfAppTitle = 262;
constexpr unsigned xlfWindowTitle = 263;
constexpr unsigned xlfSaveToolbar = 264;
constexpr unsigned xlfEnableTool = 265;
constexpr unsigned xlfPressTool = 266;
constexpr unsigned xlfRegisterId = 267;
constexpr unsigned xlfGetWorkbook = 268;
constexpr unsigned xlfAvedev = 269;
constexpr unsigned xlfBetadist = 270;
constexpr unsigned xlfGammaln = 271;
constexpr unsigned xlfBetainv = 272;
constexpr unsigned xlfBinomdist = 273;
constexpr unsigned xlfChidist = 274;
constexpr unsigned xlfChiinv = 275;
constexpr unsigned xlfCombin = 276;
constexpr unsigned xlfConfidence = 277;
constexpr unsigned xlfCritbinom = 278;
constexpr unsigned xlfEven = 279;
constexpr unsigned xlfExpondist = 280;
constexpr unsigned xlfFdist = 281;
constexpr unsigned xlfFinv = 282;
constexpr unsigned xlfFisher = 283;
constexpr unsigned xlfFisherinv = 284;
constexpr unsigned xlfFloor = 285;
constexpr unsigned xlfGammadist = 286;
constexpr unsigned xlfGammainv = 287;
constexpr unsigned xlfCeiling = 288;
constexpr unsigned xlfHypgeomdist = 289;
constexpr unsigned xlfLognormdist = 290;
constexpr unsigned xlfLoginv = 291;
constexpr unsigned xlfNegbinomdist = 292;
constexpr unsigned xlfNormdist = 293;
constexpr unsigned xlfNormsdist = 294;
constexpr unsigned xlfNorminv = 295;
constexpr unsigned xlfNormsinv = 296;
constexpr unsigned xlfStandardize = 297;
constexpr unsigned xlfOdd = 298;
constexpr unsigned xlfPermut = 299;
constexpr unsigned xlfPoisson = 300;
constexpr unsigned xlfTdist = 301;
constexpr unsigned xlfWeibull = 302;
constexpr unsigned xlfSumxmy2 = 303;
constexpr unsigned xlfSumx2my2 = 304;
constexpr unsigned xlfSumx2py2 = 305;
constexpr unsigned xlfChitest = 306;
constexpr unsigned xlfCorrel = 307;
constexpr unsigned xlfCovar = 308;
constexpr unsigned xlfForecast = 309;
constexpr unsigned xlfFtest = 310;
constexpr unsigned xlfIntercept = 311;
constexpr unsigned xlfPearson = 312;
constexpr unsigned xlfRsq = 313;
constexpr unsigned xlfSteyx = 314;
constexpr unsigned xlfSlope = 315;
constexpr unsigned xlfTtest = 316;
constexpr unsigned xlfProb = 317;
constexpr unsigned xlfDevsq = 318;
constexpr unsigned xlfGeomean = 319;
constexpr unsigned xlfHarmean = 320;
constexpr unsigned xlfSumsq = 321;
constexpr unsigned xlfKurt = 322;
constexpr unsigned xlfSkew = 323;
constexpr unsigned xlfZtest = 324;
constexpr unsigned xlfLarge = 325;
constexpr unsigned xlfSmall = 326;
constexpr unsigned xlfQuartile = 327;
constexpr unsigned xlfPercentile = 328;
constexpr unsigned xlfPercentrank = 329;
constexpr unsigned xlfMode = 330;
constexpr unsigned xlfTrimmean = 331;
constexpr unsigned xlfTinv = 332;
constexpr unsigned xlfMovieCommand = 334;
constexpr unsigned xlfGetMovie = 335;
constexpr unsigned xlfConcatenate = 336;
constexpr unsigned xlfPower = 337;
constexpr unsigned xlfPivotAddData = 338;
constexpr unsigned xlfGetPivotTable = 339;
constexpr unsigned xlfGetPivotField = 340;
constexpr unsigned xlfGetPivotItem = 341;
constexpr unsigned xlfRadians = 342;
constexpr unsigned xlfDegrees = 343;
constexpr unsigned xlfSubtotal = 344;
constexpr unsigned xlfSumif = 345;
constexpr unsigned xlfCountif = 346;
constexpr unsigned xlfCountblank = 347;
constexpr unsigned xlfScenarioGet = 348;
constexpr unsigned xlfOptionsListsGet = 349;
constexpr unsigned xlfIspmt = 350;
constexpr unsigned xlfDatedif = 351;
constexpr unsigned xlfDatestring = 352;
constexpr unsigned xlfNumberstring = 353;
constexpr unsigned xlfRoman = 354;
constexpr unsigned xlfOpenDialog = 355;
constexpr unsigned xlfSaveDialog = 356;
constexpr unsigned xlfViewGet = 357;
constexpr unsigned xlfGetpivotdata = 358;
constexpr unsigned xlfHyperlink = 359;
constexpr unsigned xlfPhonetic = 360;
constexpr unsigned xlfAveragea = 361;
constexpr unsigned xlfMaxa = 362;
constexpr unsigned xlfMina = 363;
constexpr unsigned xlfStdevpa = 364;
constexpr unsigned xlfVarpa = 365;
constexpr unsigned xlfStdeva = 366;
constexpr unsigned xlfVara = 367;
constexpr unsigned xlfBahttext = 368;
constexpr unsigned xlfThaidayofweek = 369;
constexpr unsigned xlfThaidigit = 370;
constexpr unsigned xlfThaimonthofyear = 371;
constexpr unsigned xlfThainumsound = 372;
constexpr unsigned xlfThainumstring = 373;
constexpr unsigned xlfThaistringlength = 374;
constexpr unsigned xlfIsthaidigit = 375;
constexpr unsigned xlfRoundbahtdown = 376;
constexpr unsigned xlfRoundbahtup = 377;
constexpr unsigned xlfThaiyear = 378;
constexpr unsigned xlfRtd = 379;
constexpr unsigned xlfCubevalue = 380;
constexpr unsigned xlfCubemember = 381;
constexpr unsigned xlfCubememberproperty = 382;
constexpr unsigned xlfCuberankedmember = 383;
constexpr unsigned xlfHex2bin = 384;
constexpr unsigned xlfHex2dec = 385;
constexpr unsigned xlfHex2oct = 386;
constexpr unsigned xlfDec2bin = 387;
constexpr unsigned xlfDec2hex = 388;
constexpr unsigned xlfDec2oct = 389;
constexpr unsigned xlfOct2bin = 390;
constexpr unsigned xlfOct2hex = 391;
constexpr unsigned xlfOct2dec = 392;
constexpr unsigned xlfBin2dec = 393;
constexpr unsigned xlfBin2oct = 394;
constexpr unsigned xlfBin2hex = 395;
constexpr unsigned xlfImsub = 396;
constexpr unsigned xlfImdiv = 397;
constexpr unsigned xlfImpower = 398;
constexpr unsigned xlfImabs = 399;
constexpr unsigned xlfImsqrt = 400;
constexpr unsigned xlfImln = 401;
constexpr unsigned xlfImlog2 = 402;
constexpr unsigned xlfImlog10 = 403;
constexpr unsigned xlfImsin = 404;
constexpr unsigned xlfImcos = 405;
constexpr unsigned xlfImexp = 406;
constexpr unsigned xlfImargument = 407;
constexpr unsigned xlfImconjugate = 408;
constexpr unsigned xlfImaginary = 409;
constexpr unsigned xlfImreal = 410;
constexpr unsigned xlfComplex = 411;
constexpr unsigned xlfImsum = 412;
constexpr unsigned xlfImproduct = 413;
constexpr unsigned xlfSeriessum = 414;
constexpr unsigned xlfFactdouble = 415;
constexpr unsigned xlfSqrtpi = 416;
constexpr unsigned xlfQuotient = 417;
constexpr unsigned xlfDelta = 418;
constexpr unsigned xlfGestep = 419;
constexpr unsigned xlfIseven = 420;
constexpr unsigned xlfIsodd = 421;
constexpr unsigned xlfMround = 422;
constexpr unsigned xlfErf = 423;
constexpr unsigned xlfErfc = 424;
constexpr unsigned xlfBesselj = 425;
constexpr unsigned xlfBesselk = 426;
constexpr unsigned xlfBessely = 427;
constexpr unsigned xlfBesseli = 428;
constexpr unsigned xlfXirr = 429;
constexpr unsigned xlfXnpv = 430;
constexpr unsigned xlfPricemat = 431;
constexpr unsigned xlfYieldmat = 432;
constexpr unsigned xlfIntrate = 433;
constexpr unsigned xlfReceived = 434;
constexpr unsigned xlfDisc = 435;
constexpr unsigned xlfPricedisc = 436;
constexpr unsigned xlfYielddisc = 437;
constexpr unsigned xlfTbilleq = 438;
constexpr unsigned xlfTbillprice = 439;
constexpr unsigned xlfTbillyield = 440;
constexpr unsigned xlfPrice = 441;
constexpr unsigned xlfYield = 442;
constexpr unsigned xlfDollarde = 443;
constexpr unsigned xlfDollarfr = 444;
constexpr unsigned xlfNominal = 445;
constexpr unsigned xlfEffect = 446;
constexpr unsigned xlfCumprinc = 447;
constexpr unsigned xlfCumipmt = 448;
constexpr unsigned xlfEdate = 449;
constexpr unsigned xlfEomonth = 450;
constexpr unsigned xlfYearfrac = 451;
constexpr unsigned xlfCoupdaybs = 452;
constexpr unsigned xlfCoupdays = 453;
constexpr unsigned xlfCoupdaysnc = 454;
constexpr unsigned xlfCoupncd = 455;
constexpr unsigned xlfCoupnum = 456;
constexpr unsigned xlfCouppcd = 457;
constexpr unsigned xlfDuration = 458;
constexpr unsigned xlfMduration = 459;
constexpr unsigned xlfOddlprice = 460;
constexpr unsigned xlfOddlyield = 461;
constexpr unsigned xlfOddfprice = 462;
constexpr unsigned xlfOddfyield = 463;
constexpr unsigned xlfRandbetween = 464;
constexpr unsigned xlfWeeknum = 465;
constexpr unsigned xlfAmordegrc = 466;
constexpr unsigned xlfAmorlinc = 467;
constexpr unsigned xlfConvert = 468;
constexpr unsigned xlfAccrint = 469;
constexpr unsigned xlfAccrintm = 470;
constexpr unsigned xlfWorkday = 471;
constexpr unsigned xlfNetworkdays = 472;
constexpr unsigned xlfGcd = 473;
constexpr unsigned xlfMultinomial = 474;
constexpr unsigned xlfLcm = 475;
constexpr unsigned xlfFvschedule = 476;
constexpr unsigned xlfCubekpimember = 477;
constexpr unsigned xlfCubeset = 478;
constexpr unsigned xlfCubesetcount = 479;
constexpr unsigned xlfIferror = 480;
constexpr unsigned xlfCountifs = 481;
constexpr unsigned xlfSumifs = 482;
constexpr unsigned xlfAverageif = 483;
constexpr unsigned xlfAverageifs = 484;
constexpr unsigned xlfAggregate = 485;
constexpr unsigned xlfBinom_dist = 486;
constexpr unsigned xlfBinom_inv = 487;
constexpr unsigned xlfConfidence_norm = 488;
constexpr unsigned xlfConfidence_t = 489;
constexpr unsigned xlfChisq_test = 490;
constexpr unsigned xlfF_test = 491;
constexpr unsigned xlfCovariance_p = 492;
constexpr unsigned xlfCovariance_s = 493;
constexpr unsigned xlfExpon_dist = 494;
constexpr unsigned xlfGamma_dist = 495;
constexpr unsigned xlfGamma_inv = 496;
constexpr unsigned xlfMode_mult = 497;
constexpr unsigned xlfMode_sngl = 498;
constexpr unsigned xlfNorm_dist = 499;
constexpr unsigned xlfNorm_inv = 500;
constexpr unsigned xlfPercentile_exc = 501;
constexpr unsigned xlfPercentile_inc = 502;
constexpr unsigned xlfPercentrank_exc = 503;
constexpr unsigned xlfPercentrank_inc = 504;
constexpr unsigned xlfPoisson_dist = 505;
constexpr unsigned xlfQuartile_exc = 506;
constexpr unsigned xlfQuartile_inc = 507;
constexpr unsigned xlfRank_avg = 508;
constexpr unsigned xlfRank_eq = 509;
constexpr unsigned xlfStdev_s = 510;
constexpr unsigned xlfStdev_p = 511;
constexpr unsigned xlfT_dist = 512;
constexpr unsigned xlfT_dist_2t = 513;
constexpr unsigned xlfT_dist_rt = 514;
constexpr unsigned xlfT_inv = 515;
constexpr unsigned xlfT_inv_2t = 516;
constexpr unsigned xlfVar_s = 517;
constexpr unsigned xlfVar_p = 518;
constexpr unsigned xlfWeibull_dist = 519;
constexpr unsigned xlfNetworkdays_intl = 520;
constexpr unsigned xlfWorkday_intl = 521;
constexpr unsigned xlfEcma_ceiling = 522;
constexpr unsigned xlfIso_ceiling = 523;
constexpr unsigned xlfBeta_dist = 525;
constexpr unsigned xlfBeta_inv = 526;
constexpr unsigned xlfChisq_dist = 527;
constexpr unsigned xlfChisq_dist_rt = 528;
constexpr unsigned xlfChisq_inv = 529;
constexpr unsigned xlfChisq_inv_rt = 530;
constexpr unsigned xlfF_dist = 531;
constexpr unsigned xlfF_dist_rt = 532;
constexpr unsigned xlfF_inv = 533;
constexpr unsigned xlfF_inv_rt = 534;
constexpr unsigned xlfHypgeom_dist = 535;
constexpr unsigned xlfLognorm_dist = 536;
constexpr unsigned xlfLognorm_inv = 537;
constexpr unsigned xlfNegbinom_dist = 538;
constexpr unsigned xlfNorm_s_dist = 539;
constexpr unsigned xlfNorm_s_inv = 540;
constexpr unsigned xlfT_test = 541;
constexpr unsigned xlfZ_test = 542;
constexpr unsigned xlfErf_precise = 543;
constexpr unsigned xlfErfc_precise = 544;
constexpr unsigned xlfGammaln_precise = 545;
constexpr unsigned xlfCeiling_precise = 546;
constexpr unsigned xlfFloor_precise = 547;
constexpr unsigned xlfAcot = 548;
constexpr unsigned xlfAcoth = 549;
constexpr unsigned xlfCot = 550;
constexpr unsigned xlfCoth = 551;
constexpr unsigned xlfCsc = 552;
constexpr unsigned xlfCsch = 553;
constexpr unsigned xlfSec = 554;
constexpr unsigned xlfSech = 555;
constexpr unsigned xlfImtan = 556;
constexpr unsigned xlfImcot = 557;
constexpr unsigned xlfImcsc = 558;
constexpr unsigned xlfImcsch = 559;
constexpr unsigned xlfImsec = 560;
constexpr unsigned xlfImsech = 561;
constexpr unsigned xlfBitand = 562;
constexpr unsigned xlfBitor = 563;
constexpr unsigned xlfBitxor = 564;
constexpr unsigned xlfBitlshift = 565;
constexpr unsigned xlfBitrshift = 566;
constexpr unsigned xlfPermutationa = 567;
constexpr unsigned xlfCombina = 568;
constexpr unsigned xlfXor = 569;
constexpr unsigned xlfPduration = 570;
constexpr unsigned xlfBase = 571;
constexpr unsigned xlfDecimal = 572;
constexpr unsigned xlfDays = 573;
constexpr unsigned xlfBinom_dist_range = 574;
constexpr unsigned xlfGamma = 575;
constexpr unsigned xlfSkew_p = 576;
constexpr unsigned xlfGauss = 577;
constexpr unsigned xlfPhi = 578;
constexpr unsigned xlfRri = 579;
constexpr unsigned xlfUnichar = 580;
constexpr unsigned xlfUnicode = 581;
constexpr unsigned xlfMunit = 582;
constexpr unsigned xlfArabic = 583;
constexpr unsigned xlfIsoweeknum = 584;
constexpr unsigned xlfNumbervalue = 585;
constexpr unsigned xlfSheet = 586;
constexpr unsigned xlfSheets = 587;
constexpr unsigned xlfFormulatext = 588;
constexpr unsigned xlfIsformula = 589;
constexpr unsigned xlfIfna = 590;
constexpr unsigned xlfCeiling_math = 591;
constexpr unsigned xlfFloor_math = 592;
constexpr unsigned xlfImsinh = 593;
constexpr unsigned xlfImcosh = 594;
constexpr unsigned xlfFilterxml = 595;
constexpr unsigned xlfWebservice = 596;
constexpr unsigned xlfEncodeurl = 59;

}