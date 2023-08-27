#pragma once

/* EDITED TO REMOVE DEPENDENCY ON WINDOWS.H AND ADD NAMESPACE */

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
      struct
      {
        XCHAR* data;
        bool xloil_view; // Our modification
      } str; 	/* xltypeStr */
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
        bool xloil_view; // Clearly wasn't in Microsoft's spec!
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

  constexpr int xltypeNum = 0x0001;
  constexpr int xltypeStr = 0x0002;
  constexpr int xltypeBool = 0x0004;
  constexpr int xltypeRef = 0x0008;
  constexpr int xltypeErr = 0x0010;
  constexpr int xltypeFlow = 0x0020;
  constexpr int xltypeMulti = 0x0040;
  constexpr int xltypeMissing = 0x0080;
  constexpr int xltypeNil = 0x0100;
  constexpr int xltypeSRef = 0x0400;
  constexpr int xltypeInt = 0x0800;
     
  constexpr int xlbitXLFree = 0x1000;
  constexpr int xlbitDLLFree = 0x4000;
          
  constexpr int xltypeBigData = (xltypeStr | xltypeInt);


  /*
  ** Error codes
  **
  ** Used for val.err field of XLOPER and XLOPER12 structures
  ** when constructing error XLOPERs and XLOPER12s
  */

  constexpr int xlerrNull = 0;
  constexpr int xlerrDiv0 = 7;
  constexpr int xlerrValue = 15;
  constexpr int xlerrRef = 23;
  constexpr int xlerrName = 29;
  constexpr int xlerrNum = 36;
  constexpr int xlerrNA = 42;
  constexpr int xlerrGettingData = 43;

  /*
  ** Return codes
  **
  ** These values can be returned from Excel4(), Excel4v(), Excel12() or Excel12v().
  */

  constexpr int xlretSuccess                = 0;    /* success */ 
  constexpr int xlretAbort                  = 1;    /* macro halted */
  constexpr int xlretInvXlfn                = 2;    /* invalid function number */ 
  constexpr int xlretInvCount               = 4;    /* invalid number of arguments */ 
  constexpr int xlretInvXloper              = 8;    /* invalid OPER structure */  
  constexpr int xlretStackOvfl              = 16;   /* stack overflow */  
  constexpr int xlretFailed                 = 32;   /* command failed */  
  constexpr int xlretUncalced               = 64;   /* uncalced cell */
  constexpr int xlretNotThreadSafe          = 128;  /* not allowed during multi-threaded calc */
  constexpr int xlretInvAsynchronousContext = 256;  /* invalid asynchronous function handle */
  constexpr int xlretNotClusterSafe         = 512;  /* not supported on cluster */


  /*
  ** XLL events
  **
  ** Passed in to an xlEventRegister call to register a corresponding event.
  */

  constexpr int xleventCalculationEnded = 1;   /* Fires at the end of calculation */
  constexpr int xleventCalculationCanceled = 2;    /* Fires when calculation is interrupted */


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

  constexpr int xlHpcRetSuccess = 0;
  constexpr int xlHpcRetSessionIdInvalid = -1;
  constexpr int xlHpcRetCallFailed = -2;


  /*
  ** Function number bits
  */

  constexpr int xlCommand = 0x8000;
  constexpr int xlSpecial = 0x4000;
  constexpr int xlIntl = 0x2000;
  constexpr int xlPrompt = 0x1000;

  /*
  ** Auxiliary function numbers
  **
  ** These functions are available only from the C API,
  ** not from the Excel macro language.
  */

  constexpr int xlFree = (0 | xlSpecial);
  constexpr int xlStack = (1 | xlSpecial);
  constexpr int xlCoerce = (2 | xlSpecial);
  constexpr int xlSet = (3 | xlSpecial);
  constexpr int xlSheetId = (4 | xlSpecial);
  constexpr int xlSheetNm = (5 | xlSpecial);
  constexpr int xlAbort = (6 | xlSpecial);
  constexpr int xlGetInst = (7 | xlSpecial); /* Returns application's hinstance as an integer value, supported on 32-bit platform only */
  constexpr int xlGetHwnd = (8 | xlSpecial);
  constexpr int xlGetName = (9 | xlSpecial);
  constexpr int xlEnableXLMsgs = (10 | xlSpecial);
  constexpr int xlDisableXLMsgs = (11 | xlSpecial);
  constexpr int xlDefineBinaryName = (12 | xlSpecial);
  constexpr int xlGetBinaryName = (13 | xlSpecial);
  /* GetFooInfo are valid only for calls to LPenHelper */
 constexpr int xlGetFmlaInfo	= (14 | xlSpecial);
 constexpr int xlGetMouseInfo = (15 | xlSpecial);
 constexpr int xlAsyncReturn	= (16 | xlSpecial);	/*Set return value from an asynchronous function call*/
 constexpr int xlEventRegister = (17 | xlSpecial);	/*Register an XLL event*/
 constexpr int xlRunningOnCluster = (18 | xlSpecial);	/*Returns true if running on Compute Cluster*/
 constexpr int xlGetInstPtr = (19 | xlSpecial);	/* Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms */

  /* edit modes */
  constexpr int xlModeReady = 0;	// not in edit mode
  constexpr int xlModeEnter = 1;	// enter mode
  constexpr int xlModeEdit = 2;	// edit mode
  constexpr int xlModePoint = 4;	// point mode

  /* document(page) types */
   constexpr int dtNil = 0x7f;	// window is not a sheet, macro, chart or basic
  // OR window is not the selected window at idle state
  constexpr int dtSheet = 0;// sheet
  constexpr int dtProc = 1;	// XLM macro
  constexpr int dtChart = 2;// Chart
  constexpr int dtBasic = 6;	// VBA 

   /* hit test codes */
   constexpr int htNone = 0x00;	// none of below
   constexpr int htClient = 0x01;	// internal for "in the client are", should never see
   constexpr int htVSplit = 0x02;	// vertical split area with split panes
   constexpr int htHSplit = 0x03;	// horizontal split area
   constexpr int htColWidth = 0x04;	// column width adjuster area
   constexpr int htRwHeight = 0x05;	// row height adjuster area
   constexpr int htRwColHdr = 0x06;	// the intersection of row and column headers
   constexpr int htObject = 0x07;	// the body of an object
   // the following are for size handles of draw objects
   constexpr int htTopLeft = 0x08;
   constexpr int htBotLeft = 0x09;
   constexpr int htLeft = 0x0A;
   constexpr int htTopRight = 0x0B;
   constexpr int htBotRight = 0x0C;
   constexpr int htRight = 0x0D;
   constexpr int htTop = 0x0E;
   constexpr int htBot = 0x0F;
   // end size handles
   constexpr int htRwGut = 0x10;	// row area of outline gutter
   constexpr int htColGut = 0x11;	// column area of outline gutter
   constexpr int htTextBox = 0x12;	// body of a text box (where we shouw I-Beam cursor)
   constexpr int htRwLevels = 0x13;	// row levels buttons of outline gutter
   constexpr int htColLevels = 0x14;	// column levels buttons of outline gutter
   constexpr int htDman = 0x15;	// the drag/drop handle of the selection
   constexpr int htDmanFill = 0x16;	// the auto-fill handle of the selection
   constexpr int htXSplit = 0x17;	// the intersection of the horz & vert pane splits
   constexpr int htVertex = 0x18;	// a vertex of a polygon draw object
   constexpr int htAddVtx = 0x19;	// htVertex in add a vertex mode
   constexpr int htDelVtx = 0x1A;	// htVertex in delete a vertex mode
   constexpr int htRwHdr = 0x1B;	// row header
   constexpr int htColHdr = 0x1C;	// column header
   constexpr int htRwShow = 0x1D;	// Like htRowHeight except means grow a hidden column
   constexpr int htColShow = 0x1E;	// column version of htRwShow
   constexpr int htSizing = 0x1F;	// Internal use only
   constexpr int htSxpivot = 0x20;// a drag/drop tile in a pivot table
   constexpr int htTabs = 0x21;	// the sheet paging tabs
   constexpr int htEdit = 0x22;	// Internal use only

  typedef struct _fmlainfo
  {
    int wPointMode;	// current edit mode.  0 => rest of struct undefined
    int cch;	// count of characters in formula
    char *lpch;	// poinconstexpr unsignedter to formula characters.  READ ONLY!!!
    int ichFirst;	// char offset to start of selection
    int ichLast;	// char offset to end of selection (may be > cch)
    int ichCaret;	// char offset to blinking caret
  } FMLAINFO;


  /*
  ** User defined function
  **
  ** First argument should be a function reference.
  */

constexpr int xlUDF = 255;

// Excel function numbers
constexpr int xlfCount = 0;
constexpr int xlfIsna = 2;
constexpr int xlfIserror = 3;
constexpr int xlfSum = 4;
constexpr int xlfAverage = 5;
constexpr int xlfMin = 6;
constexpr int xlfMax = 7;
constexpr int xlfRow = 8;
constexpr int xlfColumn = 9;
constexpr int xlfNa = 10;
constexpr int xlfNpv = 11;
constexpr int xlfStdev = 12;
constexpr int xlfDollar = 13;
constexpr int xlfFixed = 14;
constexpr int xlfSin = 15;
constexpr int xlfCos = 16;
constexpr int xlfTan = 17;
constexpr int xlfAtan = 18;
constexpr int xlfPi = 19;
constexpr int xlfSqrt = 20;
constexpr int xlfExp = 21;
constexpr int xlfLn = 22;
constexpr int xlfLog10 = 23;
constexpr int xlfAbs = 24;
constexpr int xlfInt = 25;
constexpr int xlfSign = 26;
constexpr int xlfRound = 27;
constexpr int xlfLookup = 28;
constexpr int xlfIndex = 29;
constexpr int xlfRept = 30;
constexpr int xlfMid = 31;
constexpr int xlfLen = 32;
constexpr int xlfValue = 33;
constexpr int xlfTrue = 34;
constexpr int xlfFalse = 35;
constexpr int xlfAnd = 36;
constexpr int xlfOr = 37;
constexpr int xlfNot = 38;
constexpr int xlfMod = 39;
constexpr int xlfDcount = 40;
constexpr int xlfDsum = 41;
constexpr int xlfDaverage = 42;
constexpr int xlfDmin = 43;
constexpr int xlfDmax = 44;
constexpr int xlfDstdev = 45;
constexpr int xlfVar = 46;
constexpr int xlfDvar = 47;
constexpr int xlfText = 48;
constexpr int xlfLinest = 49;
constexpr int xlfTrend = 50;
constexpr int xlfLogest = 51;
constexpr int xlfGrowth = 52;
constexpr int xlfGoto = 53;
constexpr int xlfHalt = 54;
constexpr int xlfPv = 56;
constexpr int xlfFv = 57;
constexpr int xlfNper = 58;
constexpr int xlfPmt = 59;
constexpr int xlfRate = 60;
constexpr int xlfMirr = 61;
constexpr int xlfIrr = 62;
constexpr int xlfRand = 63;
constexpr int xlfMatch = 64;
constexpr int xlfDate = 65;
constexpr int xlfTime = 66;
constexpr int xlfDay = 67;
constexpr int xlfMonth = 68;
constexpr int xlfYear = 69;
constexpr int xlfWeekday = 70;
constexpr int xlfHour = 71;
constexpr int xlfMinute = 72;
constexpr int xlfSecond = 73;
constexpr int xlfNow = 74;
constexpr int xlfAreas = 75;
constexpr int xlfRows = 76;
constexpr int xlfColumns = 77;
constexpr int xlfOffset = 78;
constexpr int xlfAbsref = 79;
constexpr int xlfRelref = 80;
constexpr int xlfArgument = 81;
constexpr int xlfSearch = 82;
constexpr int xlfTranspose = 83;
constexpr int xlfError = 84;
constexpr int xlfStep = 85;
constexpr int xlfType = 86;
constexpr int xlfEcho = 87;
constexpr int xlfSetName = 88;
constexpr int xlfCaller = 89;
constexpr int xlfDeref = 90;
constexpr int xlfWindows = 91;
constexpr int xlfSeries = 92;
constexpr int xlfDocuments = 93;
constexpr int xlfActiveCell = 94;
constexpr int xlfSelection = 95;
constexpr int xlfResult = 96;
constexpr int xlfAtan2 = 97;
constexpr int xlfAsin = 98;
constexpr int xlfAcos = 99;
constexpr int xlfChoose = 100;
constexpr int xlfHlookup = 101;
constexpr int xlfVlookup = 102;
constexpr int xlfLinks = 103;
constexpr int xlfInput = 104;
constexpr int xlfIsref = 105;
constexpr int xlfGetFormula = 106;
constexpr int xlfGetName = 107;
constexpr int xlfSetValue = 108;
constexpr int xlfLog = 109;
constexpr int xlfExec = 110;
constexpr int xlfChar = 111;
constexpr int xlfLower = 112;
constexpr int xlfUpper = 113;
constexpr int xlfProper = 114;
constexpr int xlfLeft = 115;
constexpr int xlfRight = 116;
constexpr int xlfExact = 117;
constexpr int xlfTrim = 118;
constexpr int xlfReplace = 119;
constexpr int xlfSubstitute = 120;
constexpr int xlfCode = 121;
constexpr int xlfNames = 122;
constexpr int xlfDirectory = 123;
constexpr int xlfFind = 124;
constexpr int xlfCell = 125;
constexpr int xlfIserr = 126;
constexpr int xlfIstext = 127;
constexpr int xlfIsnumber = 128;
constexpr int xlfIsblank = 129;
constexpr int xlfT = 130;
constexpr int xlfN = 131;
constexpr int xlfFopen = 132;
constexpr int xlfFclose = 133;
constexpr int xlfFsize = 134;
constexpr int xlfFreadln = 135;
constexpr int xlfFread = 136;
constexpr int xlfFwriteln = 137;
constexpr int xlfFwrite = 138;
constexpr int xlfFpos = 139;
constexpr int xlfDatevalue = 140;
constexpr int xlfTimevalue = 141;
constexpr int xlfSln = 142;
constexpr int xlfSyd = 143;
constexpr int xlfDdb = 144;
constexpr int xlfGetDef = 145;
constexpr int xlfReftext = 146;
constexpr int xlfTextref = 147;
constexpr int xlfIndirect = 148;
constexpr int xlfRegister = 149;
constexpr int xlfCall = 150;
constexpr int xlfAddBar = 151;
constexpr int xlfAddMenu = 152;
constexpr int xlfAddCommand = 153;
constexpr int xlfEnableCommand = 154;
constexpr int xlfCheckCommand = 155;
constexpr int xlfRenameCommand = 156;
constexpr int xlfShowBar = 157;
constexpr int xlfDeleteMenu = 158;
constexpr int xlfDeleteCommand = 159;
constexpr int xlfGetChartItem = 160;
constexpr int xlfDialogBox = 161;
constexpr int xlfClean = 162;
constexpr int xlfMdeterm = 163;
constexpr int xlfMinverse = 164;
constexpr int xlfMmult = 165;
constexpr int xlfFiles = 166;
constexpr int xlfIpmt = 167;
constexpr int xlfPpmt = 168;
constexpr int xlfCounta = 169;
constexpr int xlfCancelKey = 170;
constexpr int xlfInitiate = 175;
constexpr int xlfRequest = 176;
constexpr int xlfPoke = 177;
constexpr int xlfExecute = 178;
constexpr int xlfTerminate = 179;
constexpr int xlfRestart = 180;
constexpr int xlfHelp = 181;
constexpr int xlfGetBar = 182;
constexpr int xlfProduct = 183;
constexpr int xlfFact = 184;
constexpr int xlfGetCell = 185;
constexpr int xlfGetWorkspace = 186;
constexpr int xlfGetWindow = 187;
constexpr int xlfGetDocument = 188;
constexpr int xlfDproduct = 189;
constexpr int xlfIsnontext = 190;
constexpr int xlfGetNote = 191;
constexpr int xlfNote = 192;
constexpr int xlfStdevp = 193;
constexpr int xlfVarp = 194;
constexpr int xlfDstdevp = 195;
constexpr int xlfDvarp = 196;
constexpr int xlfTrunc = 197;
constexpr int xlfIslogical = 198;
constexpr int xlfDcounta = 199;
constexpr int xlfDeleteBar = 200;
constexpr int xlfUnregister = 201;
constexpr int xlfUsdollar = 204;
constexpr int xlfFindb = 205;
constexpr int xlfSearchb = 206;
constexpr int xlfReplaceb = 207;
constexpr int xlfLeftb = 208;
constexpr int xlfRightb = 209;
constexpr int xlfMidb = 210;
constexpr int xlfLenb = 211;
constexpr int xlfRoundup = 212;
constexpr int xlfRounddown = 213;
constexpr int xlfAsc = 214;
constexpr int xlfDbcs = 215;
constexpr int xlfRank = 216;
constexpr int xlfAddress = 219;
constexpr int xlfDays360 = 220;
constexpr int xlfToday = 221;
constexpr int xlfVdb = 222;
constexpr int xlfMedian = 227;
constexpr int xlfSumproduct = 228;
constexpr int xlfSinh = 229;
constexpr int xlfCosh = 230;
constexpr int xlfTanh = 231;
constexpr int xlfAsinh = 232;
constexpr int xlfAcosh = 233;
constexpr int xlfAtanh = 234;
constexpr int xlfDget = 235;
constexpr int xlfCreateObject = 236;
constexpr int xlfVolatile = 237;
constexpr int xlfLastError = 238;
constexpr int xlfCustomUndo = 239;
constexpr int xlfCustomRepeat = 240;
constexpr int xlfFormulaConvert = 241;
constexpr int xlfGetLinkInfo = 242;
constexpr int xlfTextBox = 243;
constexpr int xlfInfo = 244;
constexpr int xlfGroup = 245;
constexpr int xlfGetObject = 246;
constexpr int xlfDb = 247;
constexpr int xlfPause = 248;
constexpr int xlfResume = 251;
constexpr int xlfFrequency = 252;
constexpr int xlfAddToolbar = 253;
constexpr int xlfDeleteToolbar = 254;
constexpr int xlfResetToolbar = 256;
constexpr int xlfEvaluate = 257;
constexpr int xlfGetToolbar = 258;
constexpr int xlfGetTool = 259;
constexpr int xlfSpellingCheck = 260;
constexpr int xlfErrorType = 261;
constexpr int xlfAppTitle = 262;
constexpr int xlfWindowTitle = 263;
constexpr int xlfSaveToolbar = 264;
constexpr int xlfEnableTool = 265;
constexpr int xlfPressTool = 266;
constexpr int xlfRegisterId = 267;
constexpr int xlfGetWorkbook = 268;
constexpr int xlfAvedev = 269;
constexpr int xlfBetadist = 270;
constexpr int xlfGammaln = 271;
constexpr int xlfBetainv = 272;
constexpr int xlfBinomdist = 273;
constexpr int xlfChidist = 274;
constexpr int xlfChiinv = 275;
constexpr int xlfCombin = 276;
constexpr int xlfConfidence = 277;
constexpr int xlfCritbinom = 278;
constexpr int xlfEven = 279;
constexpr int xlfExpondist = 280;
constexpr int xlfFdist = 281;
constexpr int xlfFinv = 282;
constexpr int xlfFisher = 283;
constexpr int xlfFisherinv = 284;
constexpr int xlfFloor = 285;
constexpr int xlfGammadist = 286;
constexpr int xlfGammainv = 287;
constexpr int xlfCeiling = 288;
constexpr int xlfHypgeomdist = 289;
constexpr int xlfLognormdist = 290;
constexpr int xlfLoginv = 291;
constexpr int xlfNegbinomdist = 292;
constexpr int xlfNormdist = 293;
constexpr int xlfNormsdist = 294;
constexpr int xlfNorminv = 295;
constexpr int xlfNormsinv = 296;
constexpr int xlfStandardize = 297;
constexpr int xlfOdd = 298;
constexpr int xlfPermut = 299;
constexpr int xlfPoisson = 300;
constexpr int xlfTdist = 301;
constexpr int xlfWeibull = 302;
constexpr int xlfSumxmy2 = 303;
constexpr int xlfSumx2my2 = 304;
constexpr int xlfSumx2py2 = 305;
constexpr int xlfChitest = 306;
constexpr int xlfCorrel = 307;
constexpr int xlfCovar = 308;
constexpr int xlfForecast = 309;
constexpr int xlfFtest = 310;
constexpr int xlfIntercept = 311;
constexpr int xlfPearson = 312;
constexpr int xlfRsq = 313;
constexpr int xlfSteyx = 314;
constexpr int xlfSlope = 315;
constexpr int xlfTtest = 316;
constexpr int xlfProb = 317;
constexpr int xlfDevsq = 318;
constexpr int xlfGeomean = 319;
constexpr int xlfHarmean = 320;
constexpr int xlfSumsq = 321;
constexpr int xlfKurt = 322;
constexpr int xlfSkew = 323;
constexpr int xlfZtest = 324;
constexpr int xlfLarge = 325;
constexpr int xlfSmall = 326;
constexpr int xlfQuartile = 327;
constexpr int xlfPercentile = 328;
constexpr int xlfPercentrank = 329;
constexpr int xlfMode = 330;
constexpr int xlfTrimmean = 331;
constexpr int xlfTinv = 332;
constexpr int xlfMovieCommand = 334;
constexpr int xlfGetMovie = 335;
constexpr int xlfConcatenate = 336;
constexpr int xlfPower = 337;
constexpr int xlfPivotAddData = 338;
constexpr int xlfGetPivotTable = 339;
constexpr int xlfGetPivotField = 340;
constexpr int xlfGetPivotItem = 341;
constexpr int xlfRadians = 342;
constexpr int xlfDegrees = 343;
constexpr int xlfSubtotal = 344;
constexpr int xlfSumif = 345;
constexpr int xlfCountif = 346;
constexpr int xlfCountblank = 347;
constexpr int xlfScenarioGet = 348;
constexpr int xlfOptionsListsGet = 349;
constexpr int xlfIspmt = 350;
constexpr int xlfDatedif = 351;
constexpr int xlfDatestring = 352;
constexpr int xlfNumberstring = 353;
constexpr int xlfRoman = 354;
constexpr int xlfOpenDialog = 355;
constexpr int xlfSaveDialog = 356;
constexpr int xlfViewGet = 357;
constexpr int xlfGetpivotdata = 358;
constexpr int xlfHyperlink = 359;
constexpr int xlfPhonetic = 360;
constexpr int xlfAveragea = 361;
constexpr int xlfMaxa = 362;
constexpr int xlfMina = 363;
constexpr int xlfStdevpa = 364;
constexpr int xlfVarpa = 365;
constexpr int xlfStdeva = 366;
constexpr int xlfVara = 367;
constexpr int xlfBahttext = 368;
constexpr int xlfThaidayofweek = 369;
constexpr int xlfThaidigit = 370;
constexpr int xlfThaimonthofyear = 371;
constexpr int xlfThainumsound = 372;
constexpr int xlfThainumstring = 373;
constexpr int xlfThaistringlength = 374;
constexpr int xlfIsthaidigit = 375;
constexpr int xlfRoundbahtdown = 376;
constexpr int xlfRoundbahtup = 377;
constexpr int xlfThaiyear = 378;
constexpr int xlfRtd = 379;
constexpr int xlfCubevalue = 380;
constexpr int xlfCubemember = 381;
constexpr int xlfCubememberproperty = 382;
constexpr int xlfCuberankedmember = 383;
constexpr int xlfHex2bin = 384;
constexpr int xlfHex2dec = 385;
constexpr int xlfHex2oct = 386;
constexpr int xlfDec2bin = 387;
constexpr int xlfDec2hex = 388;
constexpr int xlfDec2oct = 389;
constexpr int xlfOct2bin = 390;
constexpr int xlfOct2hex = 391;
constexpr int xlfOct2dec = 392;
constexpr int xlfBin2dec = 393;
constexpr int xlfBin2oct = 394;
constexpr int xlfBin2hex = 395;
constexpr int xlfImsub = 396;
constexpr int xlfImdiv = 397;
constexpr int xlfImpower = 398;
constexpr int xlfImabs = 399;
constexpr int xlfImsqrt = 400;
constexpr int xlfImln = 401;
constexpr int xlfImlog2 = 402;
constexpr int xlfImlog10 = 403;
constexpr int xlfImsin = 404;
constexpr int xlfImcos = 405;
constexpr int xlfImexp = 406;
constexpr int xlfImargument = 407;
constexpr int xlfImconjugate = 408;
constexpr int xlfImaginary = 409;
constexpr int xlfImreal = 410;
constexpr int xlfComplex = 411;
constexpr int xlfImsum = 412;
constexpr int xlfImproduct = 413;
constexpr int xlfSeriessum = 414;
constexpr int xlfFactdouble = 415;
constexpr int xlfSqrtpi = 416;
constexpr int xlfQuotient = 417;
constexpr int xlfDelta = 418;
constexpr int xlfGestep = 419;
constexpr int xlfIseven = 420;
constexpr int xlfIsodd = 421;
constexpr int xlfMround = 422;
constexpr int xlfErf = 423;
constexpr int xlfErfc = 424;
constexpr int xlfBesselj = 425;
constexpr int xlfBesselk = 426;
constexpr int xlfBessely = 427;
constexpr int xlfBesseli = 428;
constexpr int xlfXirr = 429;
constexpr int xlfXnpv = 430;
constexpr int xlfPricemat = 431;
constexpr int xlfYieldmat = 432;
constexpr int xlfIntrate = 433;
constexpr int xlfReceived = 434;
constexpr int xlfDisc = 435;
constexpr int xlfPricedisc = 436;
constexpr int xlfYielddisc = 437;
constexpr int xlfTbilleq = 438;
constexpr int xlfTbillprice = 439;
constexpr int xlfTbillyield = 440;
constexpr int xlfPrice = 441;
constexpr int xlfYield = 442;
constexpr int xlfDollarde = 443;
constexpr int xlfDollarfr = 444;
constexpr int xlfNominal = 445;
constexpr int xlfEffect = 446;
constexpr int xlfCumprinc = 447;
constexpr int xlfCumipmt = 448;
constexpr int xlfEdate = 449;
constexpr int xlfEomonth = 450;
constexpr int xlfYearfrac = 451;
constexpr int xlfCoupdaybs = 452;
constexpr int xlfCoupdays = 453;
constexpr int xlfCoupdaysnc = 454;
constexpr int xlfCoupncd = 455;
constexpr int xlfCoupnum = 456;
constexpr int xlfCouppcd = 457;
constexpr int xlfDuration = 458;
constexpr int xlfMduration = 459;
constexpr int xlfOddlprice = 460;
constexpr int xlfOddlyield = 461;
constexpr int xlfOddfprice = 462;
constexpr int xlfOddfyield = 463;
constexpr int xlfRandbetween = 464;
constexpr int xlfWeeknum = 465;
constexpr int xlfAmordegrc = 466;
constexpr int xlfAmorlinc = 467;
constexpr int xlfConvert = 468;
constexpr int xlfAccrint = 469;
constexpr int xlfAccrintm = 470;
constexpr int xlfWorkday = 471;
constexpr int xlfNetworkdays = 472;
constexpr int xlfGcd = 473;
constexpr int xlfMultinomial = 474;
constexpr int xlfLcm = 475;
constexpr int xlfFvschedule = 476;
constexpr int xlfCubekpimember = 477;
constexpr int xlfCubeset = 478;
constexpr int xlfCubesetcount = 479;
constexpr int xlfIferror = 480;
constexpr int xlfCountifs = 481;
constexpr int xlfSumifs = 482;
constexpr int xlfAverageif = 483;
constexpr int xlfAverageifs = 484;
constexpr int xlfAggregate = 485;
constexpr int xlfBinom_dist = 486;
constexpr int xlfBinom_inv = 487;
constexpr int xlfConfidence_norm = 488;
constexpr int xlfConfidence_t = 489;
constexpr int xlfChisq_test = 490;
constexpr int xlfF_test = 491;
constexpr int xlfCovariance_p = 492;
constexpr int xlfCovariance_s = 493;
constexpr int xlfExpon_dist = 494;
constexpr int xlfGamma_dist = 495;
constexpr int xlfGamma_inv = 496;
constexpr int xlfMode_mult = 497;
constexpr int xlfMode_sngl = 498;
constexpr int xlfNorm_dist = 499;
constexpr int xlfNorm_inv = 500;
constexpr int xlfPercentile_exc = 501;
constexpr int xlfPercentile_inc = 502;
constexpr int xlfPercentrank_exc = 503;
constexpr int xlfPercentrank_inc = 504;
constexpr int xlfPoisson_dist = 505;
constexpr int xlfQuartile_exc = 506;
constexpr int xlfQuartile_inc = 507;
constexpr int xlfRank_avg = 508;
constexpr int xlfRank_eq = 509;
constexpr int xlfStdev_s = 510;
constexpr int xlfStdev_p = 511;
constexpr int xlfT_dist = 512;
constexpr int xlfT_dist_2t = 513;
constexpr int xlfT_dist_rt = 514;
constexpr int xlfT_inv = 515;
constexpr int xlfT_inv_2t = 516;
constexpr int xlfVar_s = 517;
constexpr int xlfVar_p = 518;
constexpr int xlfWeibull_dist = 519;
constexpr int xlfNetworkdays_intl = 520;
constexpr int xlfWorkday_intl = 521;
constexpr int xlfEcma_ceiling = 522;
constexpr int xlfIso_ceiling = 523;
constexpr int xlfBeta_dist = 525;
constexpr int xlfBeta_inv = 526;
constexpr int xlfChisq_dist = 527;
constexpr int xlfChisq_dist_rt = 528;
constexpr int xlfChisq_inv = 529;
constexpr int xlfChisq_inv_rt = 530;
constexpr int xlfF_dist = 531;
constexpr int xlfF_dist_rt = 532;
constexpr int xlfF_inv = 533;
constexpr int xlfF_inv_rt = 534;
constexpr int xlfHypgeom_dist = 535;
constexpr int xlfLognorm_dist = 536;
constexpr int xlfLognorm_inv = 537;
constexpr int xlfNegbinom_dist = 538;
constexpr int xlfNorm_s_dist = 539;
constexpr int xlfNorm_s_inv = 540;
constexpr int xlfT_test = 541;
constexpr int xlfZ_test = 542;
constexpr int xlfErf_precise = 543;
constexpr int xlfErfc_precise = 544;
constexpr int xlfGammaln_precise = 545;
constexpr int xlfCeiling_precise = 546;
constexpr int xlfFloor_precise = 547;
constexpr int xlfAcot = 548;
constexpr int xlfAcoth = 549;
constexpr int xlfCot = 550;
constexpr int xlfCoth = 551;
constexpr int xlfCsc = 552;
constexpr int xlfCsch = 553;
constexpr int xlfSec = 554;
constexpr int xlfSech = 555;
constexpr int xlfImtan = 556;
constexpr int xlfImcot = 557;
constexpr int xlfImcsc = 558;
constexpr int xlfImcsch = 559;
constexpr int xlfImsec = 560;
constexpr int xlfImsech = 561;
constexpr int xlfBitand = 562;
constexpr int xlfBitor = 563;
constexpr int xlfBitxor = 564;
constexpr int xlfBitlshift = 565;
constexpr int xlfBitrshift = 566;
constexpr int xlfPermutationa = 567;
constexpr int xlfCombina = 568;
constexpr int xlfXor = 569;
constexpr int xlfPduration = 570;
constexpr int xlfBase = 571;
constexpr int xlfDecimal = 572;
constexpr int xlfDays = 573;
constexpr int xlfBinom_dist_range = 574;
constexpr int xlfGamma = 575;
constexpr int xlfSkew_p = 576;
constexpr int xlfGauss = 577;
constexpr int xlfPhi = 578;
constexpr int xlfRri = 579;
constexpr int xlfUnichar = 580;
constexpr int xlfUnicode = 581;
constexpr int xlfMunit = 582;
constexpr int xlfArabic = 583;
constexpr int xlfIsoweeknum = 584;
constexpr int xlfNumbervalue = 585;
constexpr int xlfSheet = 586;
constexpr int xlfSheets = 587;
constexpr int xlfFormulatext = 588;
constexpr int xlfIsformula = 589;
constexpr int xlfIfna = 590;
constexpr int xlfCeiling_math = 591;
constexpr int xlfFloor_math = 592;
constexpr int xlfImsinh = 593;
constexpr int xlfImcosh = 594;
constexpr int xlfFilterxml = 595;
constexpr int xlfWebservice = 596;
constexpr int xlfEncodeurl = 597;

/* Excel command numbers */
constexpr int xlcBeep = 0 | xlCommand;
constexpr int xlcOpen = 1 | xlCommand;
constexpr int xlcOpenLinks = 2 | xlCommand;
constexpr int xlcCloseAll = 3 | xlCommand;
constexpr int xlcSave = 4 | xlCommand;
constexpr int xlcSaveAs = 5 | xlCommand;
constexpr int xlcFileDelete = 6 | xlCommand;
constexpr int xlcPageSetup = 7 | xlCommand;
constexpr int xlcPrint = 8 | xlCommand;
constexpr int xlcPrinterSetup = 9 | xlCommand;
constexpr int xlcQuit = 10 | xlCommand;
constexpr int xlcNewWindow = 11 | xlCommand;
constexpr int xlcArrangeAll = 12 | xlCommand;
constexpr int xlcWindowSize = 13 | xlCommand;
constexpr int xlcWindowMove = 14 | xlCommand;
constexpr int xlcFull = 15 | xlCommand;
constexpr int xlcClose = 16 | xlCommand;
constexpr int xlcRun = 17 | xlCommand;
constexpr int xlcSetPrintArea = 22 | xlCommand;
constexpr int xlcSetPrintTitles = 23 | xlCommand;
constexpr int xlcSetPageBreak = 24 | xlCommand;
constexpr int xlcRemovePageBreak = 25 | xlCommand;
constexpr int xlcFont = 26 | xlCommand;
constexpr int xlcDisplay = 27 | xlCommand;
constexpr int xlcProtectDocument = 28 | xlCommand;
constexpr int xlcPrecision = 29 | xlCommand;
constexpr int xlcA1R1c1 = 30 | xlCommand;
constexpr int xlcCalculateNow = 31 | xlCommand;
constexpr int xlcCalculation = 32 | xlCommand;
constexpr int xlcDataFind = 34 | xlCommand;
constexpr int xlcExtract = 35 | xlCommand;
constexpr int xlcDataDelete = 36 | xlCommand;
constexpr int xlcSetDatabase = 37 | xlCommand;
constexpr int xlcSetCriteria = 38 | xlCommand;
constexpr int xlcSort = 39 | xlCommand;
constexpr int xlcDataSeries = 40 | xlCommand;
constexpr int xlcTable = 41 | xlCommand;
constexpr int xlcFormatNumber = 42 | xlCommand;
constexpr int xlcAlignment = 43 | xlCommand;
constexpr int xlcStyle = 44 | xlCommand;
constexpr int xlcBorder = 45 | xlCommand;
constexpr int xlcCellProtection = 46 | xlCommand;
constexpr int xlcColumnWidth = 47 | xlCommand;
constexpr int xlcUndo = 48 | xlCommand;
constexpr int xlcCut = 49 | xlCommand;
constexpr int xlcCopy = 50 | xlCommand;
constexpr int xlcPaste = 51 | xlCommand;
constexpr int xlcClear = 52 | xlCommand;
constexpr int xlcPasteSpecial = 53 | xlCommand;
constexpr int xlcEditDelete = 54 | xlCommand;
constexpr int xlcInsert = 55 | xlCommand;
constexpr int xlcFillRight = 56 | xlCommand;
constexpr int xlcFillDown = 57 | xlCommand;
constexpr int xlcDefineName = 61 | xlCommand;
constexpr int xlcCreateNames = 62 | xlCommand;
constexpr int xlcFormulaGoto = 63 | xlCommand;
constexpr int xlcFormulaFind = 64 | xlCommand;
constexpr int xlcSelectLastCell = 65 | xlCommand;
constexpr int xlcShowActiveCell = 66 | xlCommand;
constexpr int xlcGalleryArea = 67 | xlCommand;
constexpr int xlcGalleryBar = 68 | xlCommand;
constexpr int xlcGalleryColumn = 69 | xlCommand;
constexpr int xlcGalleryLine = 70 | xlCommand;
constexpr int xlcGalleryPie = 71 | xlCommand;
constexpr int xlcGalleryScatter = 72 | xlCommand;
constexpr int xlcCombination = 73 | xlCommand;
constexpr int xlcPreferred = 74 | xlCommand;
constexpr int xlcAddOverlay = 75 | xlCommand;
constexpr int xlcGridlines = 76 | xlCommand;
constexpr int xlcSetPreferred = 77 | xlCommand;
constexpr int xlcAxes = 78 | xlCommand;
constexpr int xlcLegend = 79 | xlCommand;
constexpr int xlcAttachText = 80 | xlCommand;
constexpr int xlcAddArrow = 81 | xlCommand;
constexpr int xlcSelectChart = 82 | xlCommand;
constexpr int xlcSelectPlotArea = 83 | xlCommand;
constexpr int xlcPatterns = 84 | xlCommand;
constexpr int xlcMainChart = 85 | xlCommand;
constexpr int xlcOverlay = 86 | xlCommand;
constexpr int xlcScale = 87 | xlCommand;
constexpr int xlcFormatLegend = 88 | xlCommand;
constexpr int xlcFormatText = 89 | xlCommand;
constexpr int xlcEditRepeat = 90 | xlCommand;
constexpr int xlcParse = 91 | xlCommand;
constexpr int xlcJustify = 92 | xlCommand;
constexpr int xlcHide = 93 | xlCommand;
constexpr int xlcUnhide = 94 | xlCommand;
constexpr int xlcWorkspace = 95 | xlCommand;
constexpr int xlcFormula = 96 | xlCommand;
constexpr int xlcFormulaFill = 97 | xlCommand;
constexpr int xlcFormulaArray = 98 | xlCommand;
constexpr int xlcDataFindNext = 99 | xlCommand;
constexpr int xlcDataFindPrev = 100 | xlCommand;
constexpr int xlcFormulaFindNext = 101 | xlCommand;
constexpr int xlcFormulaFindPrev = 102 | xlCommand;
constexpr int xlcActivate = 103 | xlCommand;
constexpr int xlcActivateNext = 104 | xlCommand;
constexpr int xlcActivatePrev = 105 | xlCommand;
constexpr int xlcUnlockedNext = 106 | xlCommand;
constexpr int xlcUnlockedPrev = 107 | xlCommand;
constexpr int xlcCopyPicture = 108 | xlCommand;
constexpr int xlcSelect = 109 | xlCommand;
constexpr int xlcDeleteName = 110 | xlCommand;
constexpr int xlcDeleteFormat = 111 | xlCommand;
constexpr int xlcVline = 112 | xlCommand;
constexpr int xlcHline = 113 | xlCommand;
constexpr int xlcVpage = 114 | xlCommand;
constexpr int xlcHpage = 115 | xlCommand;
constexpr int xlcVscroll = 116 | xlCommand;
constexpr int xlcHscroll = 117 | xlCommand;
constexpr int xlcAlert = 118 | xlCommand;
constexpr int xlcNew = 119 | xlCommand;
constexpr int xlcCancelCopy = 120 | xlCommand;
constexpr int xlcShowClipboard = 121 | xlCommand;
constexpr int xlcMessage = 122 | xlCommand;
constexpr int xlcPasteLink = 124 | xlCommand;
constexpr int xlcAppActivate = 125 | xlCommand;
constexpr int xlcDeleteArrow = 126 | xlCommand;
constexpr int xlcRowHeight = 127 | xlCommand;
constexpr int xlcFormatMove = 128 | xlCommand;
constexpr int xlcFormatSize = 129 | xlCommand;
constexpr int xlcFormulaReplace = 130 | xlCommand;
constexpr int xlcSendKeys = 131 | xlCommand;
constexpr int xlcSelectSpecial = 132 | xlCommand;
constexpr int xlcApplyNames = 133 | xlCommand;
constexpr int xlcReplaceFont = 134 | xlCommand;
constexpr int xlcFreezePanes = 135 | xlCommand;
constexpr int xlcShowInfo = 136 | xlCommand;
constexpr int xlcSplit = 137 | xlCommand;
constexpr int xlcOnWindow = 138 | xlCommand;
constexpr int xlcOnData = 139 | xlCommand;
constexpr int xlcDisableInput = 140 | xlCommand;
constexpr int xlcEcho = 141 | xlCommand;
constexpr int xlcOutline = 142 | xlCommand;
constexpr int xlcListNames = 143 | xlCommand;
constexpr int xlcFileClose = 144 | xlCommand;
constexpr int xlcSaveWorkbook = 145 | xlCommand;
constexpr int xlcDataForm = 146 | xlCommand;
constexpr int xlcCopyChart = 147 | xlCommand;
constexpr int xlcOnTime = 148 | xlCommand;
constexpr int xlcWait = 149 | xlCommand;
constexpr int xlcFormatFont = 150 | xlCommand;
constexpr int xlcFillUp = 151 | xlCommand;
constexpr int xlcFillLeft = 152 | xlCommand;
constexpr int xlcDeleteOverlay = 153 | xlCommand;
constexpr int xlcNote = 154 | xlCommand;
constexpr int xlcShortMenus = 155 | xlCommand;
constexpr int xlcSetUpdateStatus = 159 | xlCommand;
constexpr int xlcColorPalette = 161 | xlCommand;
constexpr int xlcDeleteStyle = 162 | xlCommand;
constexpr int xlcWindowRestore = 163 | xlCommand;
constexpr int xlcWindowMaximize = 164 | xlCommand;
constexpr int xlcError = 165 | xlCommand;
constexpr int xlcChangeLink = 166 | xlCommand;
constexpr int xlcCalculateDocument = 167 | xlCommand;
constexpr int xlcOnKey = 168 | xlCommand;
constexpr int xlcAppRestore = 169 | xlCommand;
constexpr int xlcAppMove = 170 | xlCommand;
constexpr int xlcAppSize = 171 | xlCommand;
constexpr int xlcAppMinimize = 172 | xlCommand;
constexpr int xlcAppMaximize = 173 | xlCommand;
constexpr int xlcBringToFront = 174 | xlCommand;
constexpr int xlcSendToBack = 175 | xlCommand;
constexpr int xlcMainChartType = 185 | xlCommand;
constexpr int xlcOverlayChartType = 186 | xlCommand;
constexpr int xlcSelectEnd = 187 | xlCommand;
constexpr int xlcOpenMail = 188 | xlCommand;
constexpr int xlcSendMail = 189 | xlCommand;
constexpr int xlcStandardFont = 190 | xlCommand;
constexpr int xlcConsolidate = 191 | xlCommand;
constexpr int xlcSortSpecial = 192 | xlCommand;
constexpr int xlcGallery3dArea = 193 | xlCommand;
constexpr int xlcGallery3dColumn = 194 | xlCommand;
constexpr int xlcGallery3dLine = 195 | xlCommand;
constexpr int xlcGallery3dPie = 196 | xlCommand;
constexpr int xlcView3d = 197 | xlCommand;
constexpr int xlcGoalSeek = 198 | xlCommand;
constexpr int xlcWorkgroup = 199 | xlCommand;
constexpr int xlcFillGroup = 200 | xlCommand;
constexpr int xlcUpdateLink = 201 | xlCommand;
constexpr int xlcPromote = 202 | xlCommand;
constexpr int xlcDemote = 203 | xlCommand;
constexpr int xlcShowDetail = 204 | xlCommand;
constexpr int xlcUngroup = 206 | xlCommand;
constexpr int xlcObjectProperties = 207 | xlCommand;
constexpr int xlcSaveNewObject = 208 | xlCommand;
constexpr int xlcShare = 209 | xlCommand;
constexpr int xlcShareName = 210 | xlCommand;
constexpr int xlcDuplicate = 211 | xlCommand;
constexpr int xlcApplyStyle = 212 | xlCommand;
constexpr int xlcAssignToObject = 213 | xlCommand;
constexpr int xlcObjectProtection = 214 | xlCommand;
constexpr int xlcHideObject = 215 | xlCommand;
constexpr int xlcSetExtract = 216 | xlCommand;
constexpr int xlcCreatePublisher = 217 | xlCommand;
constexpr int xlcSubscribeTo = 218 | xlCommand;
constexpr int xlcAttributes = 219 | xlCommand;
constexpr int xlcShowToolbar = 220 | xlCommand;
constexpr int xlcPrintPreview = 222 | xlCommand;
constexpr int xlcEditColor = 223 | xlCommand;
constexpr int xlcShowLevels = 224 | xlCommand;
constexpr int xlcFormatMain = 225 | xlCommand;
constexpr int xlcFormatOverlay = 226 | xlCommand;
constexpr int xlcOnRecalc = 227 | xlCommand;
constexpr int xlcEditSeries = 228 | xlCommand;
constexpr int xlcDefineStyle = 229 | xlCommand;
constexpr int xlcLinePrint = 240 | xlCommand;
constexpr int xlcEnterData = 243 | xlCommand;
constexpr int xlcGalleryRadar = 249 | xlCommand;
constexpr int xlcMergeStyles = 250 | xlCommand;
constexpr int xlcEditionOptions = 251 | xlCommand;
constexpr int xlcPastePicture = 252 | xlCommand;
constexpr int xlcPastePictureLink = 253 | xlCommand;
constexpr int xlcSpelling = 254 | xlCommand;
constexpr int xlcZoom = 256 | xlCommand;
constexpr int xlcResume = 258 | xlCommand;
constexpr int xlcInsertObject = 259 | xlCommand;
constexpr int xlcWindowMinimize = 260 | xlCommand;
constexpr int xlcSize = 261 | xlCommand;
constexpr int xlcMove = 262 | xlCommand;
constexpr int xlcSoundNote = 265 | xlCommand;
constexpr int xlcSoundPlay = 266 | xlCommand;
constexpr int xlcFormatShape = 267 | xlCommand;
constexpr int xlcExtendPolygon = 268 | xlCommand;
constexpr int xlcFormatAuto = 269 | xlCommand;
constexpr int xlcGallery3dBar = 272 | xlCommand;
constexpr int xlcGallery3dSurface = 273 | xlCommand;
constexpr int xlcFillAuto = 274 | xlCommand;
constexpr int xlcCustomizeToolbar = 276 | xlCommand;
constexpr int xlcAddTool = 277 | xlCommand;
constexpr int xlcEditObject = 278 | xlCommand;
constexpr int xlcOnDoubleclick = 279 | xlCommand;
constexpr int xlcOnEntry = 280 | xlCommand;
constexpr int xlcWorkbookAdd = 281 | xlCommand;
constexpr int xlcWorkbookMove = 282 | xlCommand;
constexpr int xlcWorkbookCopy = 283 | xlCommand;
constexpr int xlcWorkbookOptions = 284 | xlCommand;
constexpr int xlcSaveWorkspace = 285 | xlCommand;
constexpr int xlcChartWizard = 288 | xlCommand;
constexpr int xlcDeleteTool = 289 | xlCommand;
constexpr int xlcMoveTool = 290 | xlCommand;
constexpr int xlcWorkbookSelect = 291 | xlCommand;
constexpr int xlcWorkbookActivate = 292 | xlCommand;
constexpr int xlcAssignToTool = 293 | xlCommand;
constexpr int xlcCopyTool = 295 | xlCommand;
constexpr int xlcResetTool = 296 | xlCommand;
constexpr int xlcConstrainNumeric = 297 | xlCommand;
constexpr int xlcPasteTool = 298 | xlCommand;
constexpr int xlcPlacement = 300 | xlCommand;
constexpr int xlcFillWorkgroup = 301 | xlCommand;
constexpr int xlcWorkbookNew = 302 | xlCommand;
constexpr int xlcScenarioCells = 305 | xlCommand;
constexpr int xlcScenarioDelete = 306 | xlCommand;
constexpr int xlcScenarioAdd = 307 | xlCommand;
constexpr int xlcScenarioEdit = 308 | xlCommand;
constexpr int xlcScenarioShow = 309 | xlCommand;
constexpr int xlcScenarioShowNext = 310 | xlCommand;
constexpr int xlcScenarioSummary = 311 | xlCommand;
constexpr int xlcPivotTableWizard = 312 | xlCommand;
constexpr int xlcPivotFieldProperties = 313 | xlCommand;
constexpr int xlcPivotField = 314 | xlCommand;
constexpr int xlcPivotItem = 315 | xlCommand;
constexpr int xlcPivotAddFields = 316 | xlCommand;
constexpr int xlcOptionsCalculation = 318 | xlCommand;
constexpr int xlcOptionsEdit = 319 | xlCommand;
constexpr int xlcOptionsView = 320 | xlCommand;
constexpr int xlcAddinManager = 321 | xlCommand;
constexpr int xlcMenuEditor = 322 | xlCommand;
constexpr int xlcAttachToolbars = 323 | xlCommand;
constexpr int xlcVbaactivate = 324 | xlCommand;
constexpr int xlcOptionsChart = 325 | xlCommand;
constexpr int xlcVbaInsertFile = 328 | xlCommand;
constexpr int xlcVbaProcedureDefinition = 330 | xlCommand;
constexpr int xlcRoutingSlip = 336 | xlCommand;
constexpr int xlcRouteDocument = 338 | xlCommand;
constexpr int xlcMailLogon = 339 | xlCommand;
constexpr int xlcInsertPicture = 342 | xlCommand;
constexpr int xlcEditTool = 343 | xlCommand;
constexpr int xlcGalleryDoughnut = 344 | xlCommand;
constexpr int xlcChartTrend = 350 | xlCommand;
constexpr int xlcPivotItemProperties = 352 | xlCommand;
constexpr int xlcWorkbookInsert = 354 | xlCommand;
constexpr int xlcOptionsTransition = 355 | xlCommand;
constexpr int xlcOptionsGeneral = 356 | xlCommand;
constexpr int xlcFilterAdvanced = 370 | xlCommand;
constexpr int xlcMailAddMailer = 373 | xlCommand;
constexpr int xlcMailDeleteMailer = 374 | xlCommand;
constexpr int xlcMailReply = 375 | xlCommand;
constexpr int xlcMailReplyAll = 376 | xlCommand;
constexpr int xlcMailForward = 377 | xlCommand;
constexpr int xlcMailNextLetter = 378 | xlCommand;
constexpr int xlcDataLabel = 379 | xlCommand;
constexpr int xlcInsertTitle = 380 | xlCommand;
constexpr int xlcFontProperties = 381 | xlCommand;
constexpr int xlcMacroOptions = 382 | xlCommand;
constexpr int xlcWorkbookHide = 383 | xlCommand;
constexpr int xlcWorkbookUnhide = 384 | xlCommand;
constexpr int xlcWorkbookDelete = 385 | xlCommand;
constexpr int xlcWorkbookName = 386 | xlCommand;
constexpr int xlcGalleryCustom = 388 | xlCommand;
constexpr int xlcAddChartAutoformat = 390 | xlCommand;
constexpr int xlcDeleteChartAutoformat = 391 | xlCommand;
constexpr int xlcChartAddData = 392 | xlCommand;
constexpr int xlcAutoOutline = 393 | xlCommand;
constexpr int xlcTabOrder = 394 | xlCommand;
constexpr int xlcShowDialog = 395 | xlCommand;
constexpr int xlcSelectAll = 396 | xlCommand;
constexpr int xlcUngroupSheets = 397 | xlCommand;
constexpr int xlcSubtotalCreate = 398 | xlCommand;
constexpr int xlcSubtotalRemove = 399 | xlCommand;
constexpr int xlcRenameObject = 400 | xlCommand;
constexpr int xlcWorkbookScroll = 412 | xlCommand;
constexpr int xlcWorkbookNext = 413 | xlCommand;
constexpr int xlcWorkbookPrev = 414 | xlCommand;
constexpr int xlcWorkbookTabSplit = 415 | xlCommand;
constexpr int xlcFullScreen = 416 | xlCommand;
constexpr int xlcWorkbookProtect = 417 | xlCommand;
constexpr int xlcScrollbarProperties = 420 | xlCommand;
constexpr int xlcPivotShowPages = 421 | xlCommand;
constexpr int xlcTextToColumns = 422 | xlCommand;
constexpr int xlcFormatCharttype = 423 | xlCommand;
constexpr int xlcLinkFormat = 424 | xlCommand;
constexpr int xlcTracerDisplay = 425 | xlCommand;
constexpr int xlcTracerNavigate = 430 | xlCommand;
constexpr int xlcTracerClear = 431 | xlCommand;
constexpr int xlcTracerError = 432 | xlCommand;
constexpr int xlcPivotFieldGroup = 433 | xlCommand;
constexpr int xlcPivotFieldUngroup = 434 | xlCommand;
constexpr int xlcCheckboxProperties = 435 | xlCommand;
constexpr int xlcLabelProperties = 436 | xlCommand;
constexpr int xlcListboxProperties = 437 | xlCommand;
constexpr int xlcEditboxProperties = 438 | xlCommand;
constexpr int xlcPivotRefresh = 439 | xlCommand;
constexpr int xlcLinkCombo = 440 | xlCommand;
constexpr int xlcOpenText = 441 | xlCommand;
constexpr int xlcHideDialog = 442 | xlCommand;
constexpr int xlcSetDialogFocus = 443 | xlCommand;
constexpr int xlcEnableObject = 444 | xlCommand;
constexpr int xlcPushbuttonProperties = 445 | xlCommand;
constexpr int xlcSetDialogDefault = 446 | xlCommand;
constexpr int xlcFilter = 447 | xlCommand;
constexpr int xlcFilterShowAll = 448 | xlCommand;
constexpr int xlcClearOutline = 449 | xlCommand;
constexpr int xlcFunctionWizard = 450 | xlCommand;
constexpr int xlcAddListItem = 451 | xlCommand;
constexpr int xlcSetListItem = 452 | xlCommand;
constexpr int xlcRemoveListItem = 453 | xlCommand;
constexpr int xlcSelectListItem = 454 | xlCommand;
constexpr int xlcSetControlValue = 455 | xlCommand;
constexpr int xlcSaveCopyAs = 456 | xlCommand;
constexpr int xlcOptionsListsAdd = 458 | xlCommand;
constexpr int xlcOptionsListsDelete = 459 | xlCommand;
constexpr int xlcSeriesAxes = 460 | xlCommand;
constexpr int xlcSeriesX = 461 | xlCommand;
constexpr int xlcSeriesY = 462 | xlCommand;
constexpr int xlcErrorbarX = 463 | xlCommand;
constexpr int xlcErrorbarY = 464 | xlCommand;
constexpr int xlcFormatChart = 465 | xlCommand;
constexpr int xlcSeriesOrder = 466 | xlCommand;
constexpr int xlcMailLogoff = 467 | xlCommand;
constexpr int xlcClearRoutingSlip = 468 | xlCommand;
constexpr int xlcAppActivateMicrosoft = 469 | xlCommand;
constexpr int xlcMailEditMailer = 470 | xlCommand;
constexpr int xlcOnSheet = 471 | xlCommand;
constexpr int xlcStandardWidth = 472 | xlCommand;
constexpr int xlcScenarioMerge = 473 | xlCommand;
constexpr int xlcSummaryInfo = 474 | xlCommand;
constexpr int xlcFindFile = 475 | xlCommand;
constexpr int xlcActiveCellFont = 476 | xlCommand;
constexpr int xlcEnableTipwizard = 477 | xlCommand;
constexpr int xlcVbaMakeAddin = 478 | xlCommand;
constexpr int xlcInsertdatatable = 480 | xlCommand;
constexpr int xlcWorkgroupOptions = 481 | xlCommand;
constexpr int xlcMailSendMailer = 482 | xlCommand;
constexpr int xlcAutocorrect = 485 | xlCommand;
constexpr int xlcPostDocument = 489 | xlCommand;
constexpr int xlcPicklist = 491 | xlCommand;
constexpr int xlcViewShow = 493 | xlCommand;
constexpr int xlcViewDefine = 494 | xlCommand;
constexpr int xlcViewDelete = 495 | xlCommand;
constexpr int xlcSheetBackground = 509 | xlCommand;
constexpr int xlcInsertMapObject = 510 | xlCommand;
constexpr int xlcOptionsMenono = 511 | xlCommand;
constexpr int xlcNormal = 518 | xlCommand;
constexpr int xlcLayout = 519 | xlCommand;
constexpr int xlcRmPrintArea = 520 | xlCommand;
constexpr int xlcClearPrintArea = 521 | xlCommand;
constexpr int xlcAddPrintArea = 522 | xlCommand;
constexpr int xlcMoveBrk = 523 | xlCommand;
constexpr int xlcHidecurrNote = 545 | xlCommand;
constexpr int xlcHideallNotes = 546 | xlCommand;
constexpr int xlcDeleteNote = 547 | xlCommand;
constexpr int xlcTraverseNotes = 548 | xlCommand;
constexpr int xlcActivateNotes = 549 | xlCommand;
constexpr int xlcProtectRevisions = 620 | xlCommand;
constexpr int xlcUnprotectRevisions = 621 | xlCommand;
constexpr int xlcOptionsMe = 647 | xlCommand;
constexpr int xlcWebPublish = 653 | xlCommand;
constexpr int xlcNewwebquery = 667 | xlCommand;
constexpr int xlcPivotTableChart = 673 | xlCommand;
constexpr int xlcOptionsSave = 753 | xlCommand;
constexpr int xlcOptionsSpell = 755 | xlCommand;
constexpr int xlcHideallInkannots = 808 | xlCommand;
}