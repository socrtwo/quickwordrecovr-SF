//
// OCSetupHlp.iss
// --------------
//
// OpenCandy Helper Include File
//
// This file defines functions and procedures that need to
// be called from your main installer script in order to
// initialize and setup OpenCandy.
//
// Please consult the accompanying SDK documentation for
// integration details and contact partner support for
// assistance with any advanced integration needs.
//
// IMPORTANT:
// ----------
// Publishers should have no need to modify the content
// of this file. If you are modifying this file for any
// reason other than as directed by partner support
// you are probably making a mistake. Please contact
// partner support instead.
//
// Copyright (c) 2008 - 2012 SweetLabs, Inc.
//



[Code]

//--------------------------------
// OpenCandy types
//--------------------------------

#ifdef UNICODE
type OCWString = String;
type OCAString = AnsiString;
type OCTString = OCWString;
#else
type OCAString = String;
type OCTString = OCAString;
#endif

type TOCQueryFn = function ():Integer;
type TOCNotifyProc = procedure (const iNotification: Integer);



//--------------------------------
// OpenCandy definitions
//--------------------------------

// Size of strings (including terminating character)
#define OC_STR_CHARS 1024
#define OC_GUID_CHARS 33

// Alignment options for Skip and Decline controls.
// Left align to reference window
// Right align to reference window
// Align below reference window
// Align above reference window
// Use with OC_CTRL_ALIGN_BELOW_WINDOW or OC_CTRL_ALIGN_ABOVE_WINDOW to align on reference window border
// Use with OC_CTRL_ALIGN_BELOW_WINDOW or OC_CTRL_ALIGN_ABOVE_WINDOW to align on reference window border
#define OC_CTRL_ALIGN_LEFT          1
#define OC_CTRL_ALIGN_RIGHT         2
#define OC_CTRL_ALIGN_BELOW_WINDOW  4
#define OC_CTRL_ALIGN_ABOVE_WINDOW  8
#define OC_CTRL_ALIGN_VCENTER      16
#define OC_CTRL_ALIGN_ONBORDER     32

// OpenCandy window position
#ifndef OC_WND_OFFSET_X
#define OC_WND_OFFSET_X 13
#endif
#ifndef OC_WND_OFFSET_Y
#define OC_WND_OFFSET_Y 71
#endif
#ifndef OC_WND_WIDTH
#define OC_WND_WIDTH 470
#endif
#ifndef OC_WND_HEIGHT
#define OC_WND_HEIGHT 228
#endif

// Decline control attributes
#ifndef OC_CTRL_DECLINE_OFFSET_X
#define OC_CTRL_DECLINE_OFFSET_X 12
#endif
#ifndef OC_CTRL_DECLINE_OFFSET_Y
#define OC_CTRL_DECLINE_OFFSET_Y 0
#endif
#ifndef OC_CTRL_DECLINE_FONT_NAME
#define OC_CTRL_DECLINE_FONT_NAME ''
#endif
#ifndef OC_CTRL_DECLINE_FONT_SIZE
#define OC_CTRL_DECLINE_FONT_SIZE 0
#endif
#ifndef OC_CTRL_DECLINE_ALIGNMENT
#define OC_CTRL_DECLINE_ALIGNMENT OC_CTRL_ALIGN_LEFT
#endif

// Skip control attributes
#ifndef OC_CTRL_SKIP_OFFSET_X
#define OC_CTRL_SKIP_OFFSET_X 0
#endif
#ifndef OC_CTRL_SKIP_OFFSET_Y
#define OC_CTRL_SKIP_OFFSET_Y 5
#endif
#ifndef OC_CTRL_SKIP_FONT_NAME
#define OC_CTRL_SKIP_FONT_NAME ''
#endif
#ifndef OC_CTRL_SKIP_FONT_SIZE
#define OC_CTRL_SKIP_FONT_SIZE 0
#endif
#ifndef OC_CTRL_SKIP_ALIGNMENT
#define OC_CTRL_SKIP_ALIGNMENT OC_CTRL_ALIGN_LEFT | OC_CTRL_ALIGN_BELOW_WINDOW
#endif
#ifndef OC_CTRL_SKIP_UNDERLINE
#define OC_CTRL_SKIP_UNDERLINE 1
#endif
#ifndef OC_CTRL_SKIP_TEXT_FGCOLOR
#define OC_CTRL_SKIP_TEXT_FGCOLOR ''
#endif
#ifndef OC_CTRL_SKIP_TEXT_BGCOLOR
#define OC_CTRL_SKIP_TEXT_BGCOLOR ''
#endif
#ifndef OC_CTRL_SKIP_TEXT_CUSTOM
#define OC_CTRL_SKIP_TEXT_CUSTOM ''
#endif
#ifndef OC_CTRL_SKIP_LINE_ENABLE
#define OC_CTRL_SKIP_LINE_ENABLE 0
#endif
#ifndef OC_CTRL_SKIP_LINE_FGCOLOR
#define OC_CTRL_SKIP_LINE_FGCOLOR ''
#endif

// Maximum number of client instances (must be 2)
#define OC_MAX_INSTANCES 2

// Values used with OCInit2A(), OCInit2W() APIs
#define OC_INIT_SUCCESS      0
#define OC_INIT_MODE_NORMAL  0
#define OC_INIT_MODE_REMNANT 1

// Options controlling OpenCandy initialization
#define OC_INIT_PERFORM_NOW 0
#define OC_INIT_PERFORM_BYPAGEORDER 1
#define OC_INIT_PROGRESSBAR_OFF 0
#define OC_INIT_PROGRESSBAR_ON 1

// Values used with OCGetNoCandy() API
#define OC_CANDY_ENABLED  0
#define OC_CANDY_DISABLED 1

// Offer types returned by OCGetOfferType() API
#define OC_OFFER_TYPE_NORMAL   1
#define OC_OFFER_TYPE_EMBEDDED 2

// Values returned by OCGetBannerInfo() API
#define OC_OFFER_BANNER_FOUNDNEITHER     0
#define OC_OFFER_BANNER_FOUNDTITLE       1
#define OC_OFFER_BANNER_FOUNDDESCRIPTION 2
#define OC_OFFER_BANNER_FOUNDBOTH        3

// User choice indicators returned by OCGetOfferState() API
#define OC_OFFER_CHOICE_ACCEPTED  1
#define OC_OFFER_CHOICE_DECLINED  0
#define OC_OFFER_CHOICE_NOCHOICE -1

// Values used with OCCanLeaveOfferPage() API
#define OC_OFFER_LEAVEPAGE_ALLOWED    1
#define OC_OFFER_LEAVEPAGE_DISALLOWED 0

// Values used for OCGetAsyncOfferStatus() API
#define OC_OFFER_STATUS_CANOFFER_READY    0
#define OC_OFFER_STATUS_CANOFFER_NOTREADY 1
#define OC_OFFER_STATUS_QUERYING_NOTREADY 2
#define OC_OFFER_STATUS_NOOFFERSAVAILABLE 3

// Values returned by OCRunDialog() API
#define OC_OFFER_RUNDLG_FAILURE -1

// Values returned by OCLoadOpenCandyDLL() API
#define OC_LOADOCDLL_FAILURE 0

// Values used with LogDevModeMessage() API
#define OC_DEVMSG_ERROR_TRUE  1
#define OC_DEVMSG_ERROR_FALSE 0

// Values used with SetUseDefaultColorBkgrnd() API
#define OC_USEDEFAULTCOLORBKGRND_FALSE 0
#define OC_USEDEFAULTCOLORBKGRND_TRUE 1

// Values for CheckSkipAllButtonStatus() API
#define OC_CONTROL_SKIP_USED 1
#define OC_CONTROL_SKIP_UNUSED 0

// Values for CheckDeclineOfferButtonStatus() API
#define OC_CONTROL_DECLINE_USED 1
#define OC_CONTROL_DECLINE_UNUSED 0

// Values for SkipOffer() API
#define OC_SKIPPED_NOSKIP 0
#define OC_SKIPPED_OCSKIPOWNERSCREEN 1
#define OC_SKIPPED_OPENCANDYDECLINE 2
#define OC_SKIPPED_3RDPARTYSKIP 3
#define OC_SKIPPED_OCSKIPNONOWNERSCREEN 4

// Publisher responses to SkipEnable
#define OC_SKIP_ENABLED_ALWAYS 0
#define OC_SKIP_ENABLED_OCMULTIOFFERONLY 1
#define OC_SKIP_DISABLED 2

// Publisher responses to SkipQuery
#define OC_SKIPQUERY_NOTSHOWN 0
#define OC_SKIPQUERY_SHOWN_NOTACTIVATED 1
#define OC_SKIPQUERY_SHOWN_ACTIVATED 2

// Publisher receives from SkipNotification
#define OC_SKIPNOTIFICATION_SHOWN_NOTACTIVATED 0
#define OC_SKIPNOTIFICATION_SHOWN_ACTIVATED 1

// Internal skip shown
#define OC_SKIPSHOWN_UNKNOWN 0
#define OC_SKIPSHOWN_NOTSHOWN_EXTERNAL 1
#define OC_SKIPSHOWN_SHOWN_OPENCANDY 2
#define OC_SKIPSHOWN_SHOWN_EXTERNAL 3

// Internal skip states
#define OC_SKIPSTATE_NOTSKIPPED 0
#define OC_SKIPSTATE_SKIPPED_OPENCANDY 1
#define OC_SKIPSTATE_SKIPPED_EXTERNAL 2

// Internal skip display instance special values
#define OC_SKIPDISPLAY_INSTANCE_NOTSET       0
#define OC_SKIPDISPLAY_INSTANCE_DISABLEALL  -1

// Values used in the sample installer script
//
// IMPORTANT:
// Do not modify these definitions or disable the warnings below.
// If you see warnings when you compile your script you must
// modify the values you pass when calling the OpenCandyInit procedure
// (i.e. in your .iss file) before releasing your installer publicly.
#define OC_SAMPLE_PUBLISHERNAME "OpenCandy Sample"
#define OC_SAMPLE_KEY "748ad6d80864338c9c03b664839d8161"
#define OC_SAMPLE_SECRET "dfb3a60d6bfdb55c50e1ef53249f1198"

// Default keys incase publisher does not define them
#ifndef OC_STR_KEY
	#define OC_STR_KEY ''
#endif
#ifndef OC_STR_SECRET
	#define OC_STR_SECRET ''
#endif
#ifndef OC_STR_KEY2
	#define OC_STR_KEY2 ''
#endif
#ifndef OC_STR_SECRET2
	#define OC_STR_SECRET2 ''
#endif

// Compile-time checks and defaults
#if OC_STR_MY_PRODUCT_NAME == OC_SAMPLE_PUBLISHERNAME
	#pragma warning "Do not forget to change the product name from '" + OC_SAMPLE_PUBLISHERNAME + "' to your company's product name before releasing this installer."
#endif
#if OC_STR_KEY == OC_SAMPLE_KEY
	#pragma warning "Do not forget to change the sample key '" + OC_SAMPLE_KEY + "' to your company's product key before releasing this installer."
#endif
#if OC_STR_SECRET == OC_SAMPLE_SECRET
	#pragma warning "Do not forget to change the sample secret '" + OC_SAMPLE_SECRET + "' to your company's product secret before releasing this installer."
#endif
#if Pos(LowerCase("\OCSetupHlp.dll"), LowerCase(OC_OCSETUPHLP_FILE_PATH)) == 0
	#pragma error "The definition OC_OCSETUPHLP_FILE_PATH does not use ""OCSetupHlp.dll"" for the file part."
#endif

// OC_MAX_INIT_TIME_* is the maximum time in milliseconds that OCInit may block when fetching offers.
// If you intend to override these defaults, do so by defining them in your .iss file before #include'ing this header.
// Be certain to make OpenCandy partner support aware of any override you apply because this can affect your metrics.
#ifndef OC_MAX_INIT_TIME_INIT_PERFORM_NOW
	#define OC_MAX_INIT_TIME_INIT_PERFORM_NOW 8000
#endif

#ifndef OC_MAX_INIT_TIME_INIT_PERFORM_BYPAGEORDER
	#define OC_MAX_INIT_TIME_INIT_PERFORM_BYPAGEORDER 0
#endif

// OC_MAX_LOADING_TIME_* is the maximum time in milliseconds that the loading page may be displayed.
// Note that under normal network conditions the loading page may end sooner. Setting this value too low
// may reduce offer rate. Values of at least 5000 are recommended. If you intend to override these defaults do so by
// defining them in your .iss file before #include'ing this header. Be certain to make OpenCandy partner support aware
// of any override you apply because this can affect your metrics.
#ifndef OC_MAX_LOADING_TIME_INIT_PERFORM_NOW
	#define OC_MAX_LOADING_TIME_INIT_PERFORM_NOW 8000
#endif

#ifndef OC_MAX_LOADING_TIME_INIT_PERFORM_BYPAGEORDER
	#define OC_MAX_LOADING_TIME_INIT_PERFORM_BYPAGEORDER 10000
#endif


// This value controls whether a progress bar is shown while the OpenCandy SDK initializes if initialization is
// configured to block for a limited period of time. If you intend to override this default do so by defining
// it your .iss file before #include'ing this header.
#ifndef OC_INIT_PROGRESSBAR_ENABLE
	#define OC_INIT_PROGRESSBAR_ENABLE OC_INIT_PROGRESSBAR_ON
#endif

// This value prevents the progress bar from being displayed until the blocking time has exceeded the threshold.
// If you intend to override this default do so by defining it your .iss file before #include'ing this header.
#ifndef OC_INIT_PROGRESSBAR_DELAY
	#define OC_INIT_PROGRESSBAR_DELAY 4
#endif



//--------------------------------
// OpenCandy global variables
//--------------------------------

// IMPORTANT:
// ----------
// Never modify or reference these variables directly in any other script,
// they are used completely internally to this helper script and are subject
// to change without notice at any time.

var

	gl_OC_objOCLoadDLLPage: TWizardPage;                                         // Handle to the LoadDLL placeholder page
	gl_OC_objOCConnectPage: TWizardPage;                                         // Handle to the Init2 placeholder page
	gl_OC_objLoadingPage: TWizardPage;                                           // Handle to the Loading placeholder page
	gl_OC_objOCOfferPage: array [1..{#OC_MAX_INSTANCES}] of TWizardPage;         // Handle to the offer page wizard page
	gl_OC_init_tszProductName: OCTString;                                        // Product name passed to Init procedure
	gl_OC_init_tszProductKey: array [1..{#OC_MAX_INSTANCES}] of OCTString;       // Product key passed to Init procedure
	gl_OC_init_tszProductSecret: array [1..{#OC_MAX_INSTANCES}] of OCTString;    // Product secret passed to Init procedure
	gl_OC_init_tszInstallerLanguage: OCTString;                                  // Installer language passed to Init procedure
	gl_OC_init_iInitModeOffer: array [1..{#OC_MAX_INSTANCES}] of Integer;        // Init mode passed to Init procedure
	gl_OC_init_iPerformInit: Integer;                                            // When should initialization be performed?
	gl_OC_bAttached: array [1..{#OC_MAX_INSTANCES}] of Boolean;                  // Is the OpenCandy offer window attached?
	gl_OC_bHasDone_LoadDLL: Boolean;                                             // Has the Load DLL procedure been done?
	gl_OC_bHasDone_Init: Boolean;                                                // Has the client init procedure been done?
	gl_OC_bHasBeenInitialized: array [1..{#OC_MAX_INSTANCES}] of Boolean;        // Has the OpenCandy client DLL been initialized?
	gl_OC_bNoCandy: Boolean;                                                     // Is OpenCandy disabled?
	gl_OC_bOfferIsEnabled: array [1..{#OC_MAX_INSTANCES}] of Boolean;            // Is the OpenCandy offer enabled?
	gl_OC_bUseOfferPage: array [1..{#OC_MAX_INSTANCES}] of Boolean;              // Should show an offer?
	gl_OC_bHasReachedOCPage: array [1..{#OC_MAX_INSTANCES}] of Boolean;          // Has the user reached the OpenCandy offer page?
	gl_OC_bHasDone_ExecOfferEmbedded: array [1..{#OC_MAX_INSTANCES}] of Boolean; // Has the embedded mode install been handled?
	gl_OC_bProductInstallSuccess: Boolean;                                       // Has the publisher product install completed successfully?
	gl_OC_szLoadingMsg: OCTString;                                               // Custom loading message
	gl_OC_szFontName: OCTString;                                                 // Custom loading font name
	gl_OC_iFontSize: Integer;                                                    // Custom loading font size
	gl_OC_iLoadingScreenInstance: Integer;                                       // Instance ID showing loading screen
	gl_OC_bHasDone_LoadingScreen: Boolean;                                       // Have the loading screen already been shown?
	gl_OC_bAutoNextAfterLoadingScreen: Boolean;                                  // Can we skip to next page after loading screen was shown?
	gl_OC_bUseDefaultColorBkGround: array [1..{#OC_MAX_INSTANCES}] of Boolean;   // Use Windows COLOR_3DFACE system color for background of loading and offer screen 1?
	gl_OC_szCustomBrushColor: array [1..{#OC_MAX_INSTANCES}] of OCTString;       // If not empty string, custom color for background of offer screen as '#RGB' where R, G, B are hex in range 00-FF.
	gl_OC_szCustomImagePath: array [1..{#OC_MAX_INSTANCES}] of OCTString;        // If not empty string, custom image to draw as the background for offer screen, e.g. 'c:\temp\bkground.png'
	gl_OC_bHasDone_ssDone: Boolean;                                              // Has the ssDone step already been reached? (Workaround for unexpected Inno Setup behavior)
	gl_OC_szAPIGuid: array [1..{#OC_MAX_INSTANCES}] of OCAString;                // API identifier
	gl_OC_szMemMapGuid: array [1..{#OC_MAX_INSTANCES}] of OCAString;             // Mem map identifier
	gl_OC_iSkipState: Integer;                                                   // Skip state
	gl_OC_iSkipShown: Integer;                                                   // Has Skip control been shown?
	gl_OC_iSkipDisplayInstance: Integer;                                         // Instance associated with Skip control
	gl_OC_fnCallbackSkipQuery: TOCQueryFn;                                       // SkipQuery callback function
	gl_OC_fnCallbackSkipEnable: TOCQueryFn;                                      // SkipEnable callback function
	gl_OC_fnCallbackSkipNotify: TOCNotifyProc;                                   // SkipNotify callback procedure



//-----------------------------------------
// OpenCandy external procedure definitions
//-----------------------------------------

procedure _OCDLL_OCStartDLMgr2Download(pszGuid:OCAString);
external 'OCPID720OpenCandy29@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCLoadOpenCandyDLL(pszGuid:OCAString; nBufferSize:Integer; pszMemMap:OCAString; nBufferSizeMemMap:Integer; pszClientSessionId:OCAString; nBufferSizeSession:Integer):Integer;
external 'OCPID720OpenCandy1@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

#ifdef UNICODE
function _OCDLL_OCInitW(pszGuid:OCAString; wszPubId, wszProdId, wszSecret, wszInstallLang:OCWString; bAsyncMode:Boolean; iMaxWait, iRemnant, iShowLoadingProgBar, iShowLoadingProgBarDelay:Integer):Integer;
external 'OCPID720OpenCandy6@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
function _OCDLL_OCInit(pszGuid, szPubId, szProdId, szSecret, szInstallLang:OCAString; bAsyncMode:Boolean; iMaxWait, iRemnant, iShowLoadingProgBar, iShowLoadingProgBarDelay:Integer):Integer;
external 'OCPID720OpenCandy5@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

#ifdef UNICODE
function _OCDLL_OCGetBannerInfoW(pszGuid, pszMemMap:OCAString; wszTitle, wszDesc:OCWString):Integer;
external 'OCPID720OpenCandy8@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
function _OCDLL_OCGetBannerInfo(pszGuid, pszMemMap, szTitle, szDesc:OCAString):Integer;
external 'OCPID720OpenCandy7@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

function _OCDLL_OCRunDialog(pszGuid:OCAString; iHwnd:Integer):Integer;
external 'OCPID720OpenCandy9@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCAdjustPage(pszGuid:OCAString; iHwnd, iX, iY, iW, iH:Integer):Integer;
external 'OCPID720OpenCandy13@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCRestorePage(pszGuid:OCAString; iHwnd:Integer):Integer;
external 'OCPID720OpenCandy14@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCGetOfferState(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy10@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCGetOfferType(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy17@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCPrepareDownload(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy18@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCShutdown(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy11@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCDetach(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy12@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCSignalProductInstalled(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy19@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCSignalProductFailed(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy20@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCGetAsyncOfferStatus(pszGuid:OCAString; bWantToShowOffer:Boolean):Integer;
external 'OCPID720OpenCandy31@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCCanLeaveOfferPage(pszGuid, pszMemMap:OCAString):Integer;
external 'OCPID720OpenCandy34@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

#ifdef UNICODE
function _OCDLL_OCSetCmdLineValuesW(pszGuid:OCAString; wszValue:OCWString):Integer;
external 'OCPID720OpenCandy36@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
function _OCDLL_OCSetCmdLineValues(pszGuid, szValue:OCAString):Integer;
external 'OCPID720OpenCandy35@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

function _OCDLL_OCGetNoCandy(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy32@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCSetNoCandy(pszGuid:OCAString; bNoCandy:Boolean):Integer;
external 'OCPID720OpenCandy33@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

procedure _OCDLL_SetOCOfferEnabled(pszGuid:OCAString; bEnabled:Boolean);
external 'OCPID720OpenCandy37@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

#ifdef UNICODE
procedure _OCDLL_LogDevModeMessageW(pszGuid:OCAString; wszMessage:OCWString; iError, iFaqID:Integer);
external 'OCPID720OpenCandy39@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
procedure _OCDLL_LogDevModeMessage(pszGuid, szMessage:OCAString; iError, iFaqID:Integer);
external 'OCPID720OpenCandy38@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

#ifdef UNICODE
procedure _OCDLL_SetClientAdvancedOptionsW(pszGuid:OCAString; szOptions:OCWString);
external 'OCPID720OpenCandy48@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
procedure _OCDLL_SetClientAdvancedOptions(pszGuid, szOptions:OCAString);
external 'OCPID720OpenCandy47@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

#ifdef UNICODE
function _OCKernel32DLL_GetModuleHandleW(szModuleName:OCWString):Integer;
external 'GetModuleHandleW@kernel32.dll stdcall';
#else
function _OCKernel32DLL_GetModuleHandle(szModuleName:OCAString):Integer;
external 'GetModuleHandleA@kernel32.dll stdcall';
#endif

procedure _OCDLL_OCShowLoadingScreen(pszGuid:OCAString; maxWaitSecs, iHwnd:Integer);
external 'OCPID720OpenCandy50@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

#ifdef UNICODE
procedure _OCDLL_OCShowLoadingScreen2W(pszGuid:OCAString; maxWaitSecs, iHwnd:Integer; wszLoadingMsg, wszFontName:OCWString; iFontSize:Integer);
external 'OCPID720OpenCandy52@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
procedure _OCDLL_OCShowLoadingScreen2A(pszGuid:OCAString; maxWaitSecs, iHwnd:Integer; wszLoadingMsg, wszFontName:OCAString; iFontSize:Integer);
external 'OCPID720OpenCandy51@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

procedure _OCDLL_OCHideLoadingScreen(pszGuid:OCAString);
external 'OCPID720OpenCandy53@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

procedure _OCDLL_OCSetUseDefaultColorBkgrnd(pszGuid:OCAString; bUseDefault:Integer);
external 'OCPID720OpenCandy44@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

#ifdef UNICODE
procedure _OCDLL_OCSetCustomBrushColorW(pszGuid:OCAString; wszColorValue:OCWString);
external 'OCPID720OpenCandy46@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
procedure _OCDLL_OCSetCustomBrushColorA(pszGuid, szColorValue:OCAString);
external 'OCPID720OpenCandy45@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

#ifdef UNICODE
procedure _OCDLL_OCSetBkGrdImagePathW(pszGuid:OCAString; wszImagePath:OCWString);
external 'OCPID720OpenCandy55@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
procedure _OCDLL_OCSetBkGrdImagePathA(pszGuid, szImagePath:OCAString);
external 'OCPID720OpenCandy54@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

#ifdef UNICODE
procedure _OCDLL_OCShowSkipAllButtonW(pszGuid:OCAString; hwndLinkPlaceHolder:LongInt; nSubClassType, nXOffSet, nYOffSet:Integer; hWndNextButton:LongInt; nDrawLine: Integer; pszLineColor, pszFontName:OCWString; nFontSize, nUseUnderlineFont:Integer; pszTextColor, pszBkGroundColor, pszCustomText:OCWString);
external 'OCPID720OpenCandy62@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#else
procedure _OCDLL_OCShowSkipAllButtonA(pszGuid:OCAString; hwndLinkPlaceHolder:LongInt; nSubClassType, nXOffSet, nYOffSet:Integer; hWndNextButton:LongInt; nDrawLine: Integer; pszLineColor, pszFontName:OCAString; nFontSize, nUseUnderlineFont:Integer; pszTextColor, pszBkGroundColor, pszCustomText:OCAString);
external 'OCPID720OpenCandy56@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';
#endif

procedure _OCDLL_OCSetDeclineButton(pszGuid:OCAString; hWndSubClass:LongInt; nSubClassType, xOffset, yOffset:Integer; hWndNextButton:LongInt; pszFontName:OCAString; nFontSize:Integer);
external 'OCPID720OpenCandy60@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCWasSkipOffersClicked(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy58@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

function _OCDLL_OCWasDeclineOfferClicked(pszGuid:OCAString):Integer;
external 'OCPID720OpenCandy61@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';

procedure _OCDLL_OCSkipOffer(pszGuid:OCAString; nSigal:Integer);
external 'OCPID720OpenCandy59@{tmp}\OCSetupHlp.dll cdecl loadwithalteredsearchpath delayload';



//-------------------------------------------
// OpenCandy runtime functions and procedures
//-------------------------------------------

//
// _OpenCandyPrepareInnoAPI
// ------------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// Prepares the OpenCandy API by initializing state for the Inno Setup
// interface layer.
//

procedure _OpenCandyPrepareInnoAPI();
var
	iOCInstance: Integer;
begin
	// Initialize OpenCandy variables
	gl_OC_objOCConnectPage := nil;
	gl_OC_objLoadingPage := nil;
	gl_OC_init_tszProductName := '';
	gl_OC_init_tszInstallerLanguage := '';
	gl_OC_init_iPerformInit := {#OC_INIT_PERFORM_NOW};
	gl_OC_bHasDone_LoadDLL := false;
	gl_OC_bHasDone_Init := false;
	gl_OC_bNoCandy := false;
	gl_OC_bProductInstallSuccess := false;
	gl_OC_szLoadingMsg := 'Loading...';
	gl_OC_szFontName := 'Arial';
	gl_OC_iFontSize := 100;
	gl_OC_iLoadingScreenInstance := 0;
	gl_OC_bHasDone_LoadingScreen := false;
	gl_OC_bAutoNextAfterLoadingScreen := false;
	gl_OC_bHasDone_ssDone := false;
	gl_OC_iSkipState := {#OC_SKIPSTATE_NOTSKIPPED}
	gl_OC_iSkipShown := {#OC_SKIPSHOWN_UNKNOWN}
	gl_OC_iSkipDisplayInstance := {#OC_SKIPDISPLAY_INSTANCE_NOTSET}
	gl_OC_fnCallbackSkipQuery := nil;
	gl_OC_fnCallbackSkipEnable := nil;
	gl_OC_fnCallbackSkipNotify := nil;
	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
	begin
		gl_OC_objOCOfferPage[iOCInstance] := nil;
		gl_OC_init_tszProductSecret[iOCInstance] := '';
		gl_OC_init_tszProductKey[iOCInstance] := '';
		gl_OC_init_iInitModeOffer[iOCInstance] := {#OC_INIT_MODE_NORMAL};
		gl_OC_bHasBeenInitialized[iOCInstance] := false;
		gl_OC_bAttached[iOCInstance] := false;
		gl_OC_bUseOfferPage[iOCInstance] := false;
		gl_OC_bHasReachedOCPage[iOCInstance] := false;
		gl_OC_bHasDone_ExecOfferEmbedded[iOCInstance] := false;
		gl_OC_bOfferIsEnabled[iOCInstance] := true;
		gl_OC_bUseDefaultColorBkGround[iOCInstance] := true;
		gl_OC_szCustomBrushColor[iOCInstance] := '';
		gl_OC_szCustomImagePath[iOCInstance] := '';
		gl_OC_szAPIGuid[iOCInstance] := '';
		gl_OC_szMemMapGuid[iOCInstance] := '';
	end;
end;



//
// _OCValidInstance
// --------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// Parameters:
//
//   iOCInstance : Client instance ID
//

function _OCValidInstance(iOCInstance:Integer):Boolean;
begin
	Result := (iOCInstance >= 1) and (iOCInstance <= {#OC_MAX_INSTANCES});
end;



//
// _OpenCandyDevModeMsg
// --------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// Parameters:
//
//   iOCInstance : Client instance ID
//   tszMessage  : Message to display
//   bIsError    : True if the message represents an error
//   iFaqID      : ID of the FAQ associated with the message, or 0 if there is no FAQ associated
//
// Usage:
//
//   _OpenCandyDevModeMsg('This is an error with associated FAQ #500', true, 500);
//

procedure _OpenCandyDevModeMsg_Instance(iOCInstance:Integer; tszMessage:OCTString; bIsError:Boolean; iFaqID:Integer);
var
	iError:Integer;
begin
	if not _OCValidInstance(iOCInstance) then
		Exit;
	if (gl_OC_szAPIGuid[iOCInstance] = '') or gl_OC_bNoCandy then
		Exit;

	if bIsError then
	begin
		iError := {#OC_DEVMSG_ERROR_TRUE};
		tszMessage := '{\rtf1 {\colortbl;\red0\green0\blue0;\red255\green0\blue0;}\cf2Status ERROR! \cf1' + tszMessage + '\par}';
	end
	else
		iError := {#OC_DEVMSG_ERROR_FALSE};

	 #ifdef UNICODE
	 _OCDLL_LogDevModeMessageW(gl_OC_szAPIGuid[iOCInstance], tszMessage, iError, iFaqID);
	 #else
	 _OCDLL_LogDevModeMessage(gl_OC_szAPIGuid[iOCInstance], tszMessage, iError, iFaqID);
	 #endif
end;



//
// _OCSetNoCandy
// -------------
// This function is internal to this helper script. Do not
// call it from your own code.
//

procedure _OCSetNoCandy();
var
	iOCInstance: Integer;
begin
	if not gl_OC_bNoCandy then
		for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
			if gl_OC_szAPIGuid[iOCInstance] <> '' then
				_OCDLL_OCSetNoCandy(gl_OC_szAPIGuid[iOCInstance], true);
	gl_OC_bNoCandy := true;
end;



//
// _OpenCandyAutoSelfDisable
// -------------------------
// This procedure is internal to this helper script. Do not
// call it from your own code.
//

procedure _OpenCandyAutoSelfDisable();
var
	i:Integer;
begin
	if not gl_OC_bNoCandy then
	begin
		// OpenCandy is disabled during a silent installation
		if WizardSilent() then
		begin
			_OCSetNoCandy();
			Exit;
		end;

		// OpenCandy may be explicitly disabled via command argument
		for i := 0 to ParamCount() do
		begin
			if AnsiUppercase(Trim(ParamStr(i))) = '/NOCANDY' then
				_OCSetNoCandy();
		end;
	end;
end;



// OCRegisterSkipQueryCallback
// --------------------------
// Registers a callback that will inform OpenCandy about whether a Skip control was presented to
// the end user before OpenCandy appeared in the wizard sequence, and what the result of that
// presentation was.
//
// Note that this function may be called multiple times during a session as the end user attempts
// to navigate back and forth between pages.
//
// The callback function must have the following prototype:
//
//   function ():Integer;
//
// The callback function must return one of these values:
//
//   OC_SKIPQUERY_SHOWN_NOTACTIVATED : A Skip control was shown but the user didn't use it.
//   OC_SKIPQUERY_SHOWN_ACTIVATED    : A Skip control was shown and the user did use it.
//   OC_SKIPQUERY_NOTSHOWN           : No Skip control was shown by an earlier non-OpenCandy offer page.
//

procedure OCRegisterSkipQueryCallback(const fnCallback: TOCQueryFn);
begin
	gl_OC_fnCallbackSkipQuery := fnCallback;
end;



// OCRegisterSkipEnableCallback
// ---------------------------
// Registers a callback that will inform OpenCandy about whether a Skip control should be shown
// on the OpenCandy offer page, and under what circumstances.
//
// The callback function must have the following prototype:
//
//   function ():Integer;
//
// The callback function must return one of these values:
//
//   OC_SKIP_ENABLED_OCMULTIOFFERONLY : Show a Skip control only if it is possible that
//                                      OpenCandy may display multiple offers.
//   OC_SKIP_ENABLED_ALWAYS           : Always show a Skip control. Use this value if
//                                      you have any other offer after the OpenCandy
//                                      offer screen.
//   OC_SKIP_DISABLED                 : Do not show a Skip control.
//

procedure OCRegisterSkipEnableCallback(const fnCallback: TOCQueryFn);
begin
	gl_OC_fnCallbackSkipEnable := fnCallback;
end;



// OCRegisterSkipNotifyCallback
// ---------------------------
// Registers a callback that receives notifications after OpenCandy shows a Skip control, including
// whether the control was activated by the end user.
//
// Note that this function may be called multiple times during a session as the end user attempts
// to navigate back and forth between pages.
//
// The callback procedure must have the following prototype:
//
//   procedure (iSkipNotification: Integer);
//
// Possible values for iSkipNotification are:
//
//   OC_SKIPNOTIFICATION_SHOWN_NOTACTIVATED : A Skip control was shown but wasn't activated.
//   OC_SKIPNOTIFICATION_SHOWN_ACTIVATED    : A Skip control was shown and was activated.
//

procedure OCRegisterSkipNotifyCallback(const procCallback: TOCNotifyProc);
begin
	gl_OC_fnCallbackSkipNotify := procCallback;
end;



//
// _OpenCandyLoadDLL
// -----------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// Loads the OpenCandy Network Client DLL. This function must be called
// before any stateful operations. The result will be true on success.
//
// Usage:
//
//   _OpenCandyLoadDLL();
//

procedure _OpenCandyLoadDLL();
var
	iRes:Integer;
	hMod:Integer;
	szAPIGuid: OCAString;
	szMemMapGuid: OCAString;
	szClientSessionId: OCAString;
	cchClientSessionId: Integer;
	iOCInstance: Integer;
begin
	if gl_OC_bNoCandy or gl_OC_bHasDone_LoadDLL then
		Exit;

	gl_OC_bHasDone_LoadDLL := true;

	// Prevent loading multiple instances
	#ifdef UNICODE
	hMod := _OCKernel32DLL_GetModuleHandleW('OCSetupHlp.dll')
	#else
	hMod := _OCKernel32DLL_GetModuleHandle('OCSetupHlp.dll')
	#endif
	if (0 <> hMod) or FileExists(ExpandConstant('{tmp}\OCSetupHlp.dll')) then
		Exit;

	// Extract and load OpenCandy Network Client library
	try
		ExtractTemporaryFile('OCSetupHlp.dll');
		cchClientSessionId := {#OC_GUID_CHARS};
		szClientSessionId := StringOfChar(#0, cchClientSessionId);
		for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
			if (gl_OC_init_tszProductKey[iOCInstance] <> '') and not gl_OC_bNoCandy then
			begin
				szAPIGuid := StringOfChar(' ', {#OC_GUID_CHARS});
				szMemMapGuid := StringOfChar(' ', {#OC_GUID_CHARS});
				iRes := _OCDLL_OCLoadOpenCandyDLL(szAPIGuid, {#OC_GUID_CHARS}, szMemMapGuid, {#OC_GUID_CHARS}, szClientSessionId, cchClientSessionId);
				if {#OC_LOADOCDLL_FAILURE} <> iRes then
				begin
					gl_OC_szAPIGuid[iOCInstance] := szAPIGuid;
					gl_OC_szMemMapGuid[iOCInstance] := szMemMapGuid;
					cchClientSessionId := Length(szClientSessionId) + 1;
					_OpenCandyDevModeMsg_Instance(iOCInstance, 'Client for product key ' + IntToStr(iOCInstance) + ' loaded.', false, 0);
				end
				else
					_OCSetNoCandy();
			end;
	except
		iRes := {#OC_LOADOCDLL_FAILURE};
	end;

	if {#OC_LOADOCDLL_FAILURE} = iRes then
		_OCSetNoCandy();
end;



//
// _OCEnabledAndReady
// ------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// Returns true if the SDK is enabled and client has been initialized.
//
// Parameters:
//
//   iOCInstance : Client instance ID to test for initialization, or
//                 zero to accept any initialized SDK.
//

function _OCEnabledAndReady(iOCInstance:Integer):Boolean;
begin
	Result := false;
	if gl_OC_bNoCandy then
		Exit;

	if (0 <> iOCInstance) then
		Result := gl_OC_bHasBeenInitialized[iOCInstance]
	else
		for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
			if gl_OC_bHasBeenInitialized[iOCInstance] then
				Result := true;
end;



//
// SetOCOfferEnabled
// -----------------
// Allows you to disable one or both OpenCandy offer screens easily from your
// installer code without affecting product installation analytics.
// This must be performed only before an offer is shown.
// Note that this is not the recommended method - you ought to determine
// during initialization whether OpenCandy should be disabled and specify
// an appropriate mode when calling OpenCandyInit or OpenCandyAsyncInit
// in that case. If you must use this method please be sure to inform the
// OpenCandy partner support team. Never directly place logical conditions
// around other OpenCandy functions and macros because this can have
// unforeseen consequences. You should call this procedure only after
// calling OpenCandyInit.
//
// Parameters:
//
//   iOCInstance : Client instance ID
//   bEnabled    : Should the offer be enabled?
//
// Usage:
//
//  // This turns off OpenCandy offer screen one when called before the offer screen
//  SetOCOfferEnabled(1, false);
//

procedure SetOCOfferEnabled(iOCInstance:Integer; bEnabled:Boolean);
begin
	if not _OCValidInstance(iOCInstance) then
		Exit;

	if _OCEnabledAndReady(iOCInstance) then
		_OCDLL_SetOCOfferEnabled(gl_OC_szAPIGuid[iOCInstance], bEnabled);
	gl_OC_bOfferIsEnabled[iOCInstance] := bEnabled;
end;



//
// OpenCandySetUseDefaultColorBkgrnd
// ---------------------------------
// Calling this procedure after OpenCandyInit in the Inno Setup
// InitializeWizard callback procedure tells the client whether to
// draw the OpenCandy loading and offer screens on the solid
// Windows system color COLOR_3DFACE. The OpenCandy loading screen
// takes the same configuration as the first enabled, non-remnant
// OpenCandy offer screen.
//
// Parameters:
//
//   iOCInstance : Client instance ID
//   bUseDefault : Use the default solid background color?
//
// Usage:
//
//   // Do not use default solid background color for OpenCandy offer screen one
//   OpenCandySetUseDefaultColorBkgrnd(1, false);
//

procedure OpenCandySetUseDefaultColorBkgrnd(iOCInstance:Integer; bUseDefault:Boolean);
begin
	if not _OCValidInstance(iOCInstance) then
		Exit;

	if _OCEnabledAndReady(iOCInstance) then
		if (bUseDefault) then
			_OCDLL_OCSetUseDefaultColorBkgrnd(gl_OC_szAPIGuid[iOCInstance], {#OC_USEDEFAULTCOLORBKGRND_TRUE})
		else
			_OCDLL_OCSetUseDefaultColorBkgrnd(gl_OC_szAPIGuid[iOCInstance], {#OC_USEDEFAULTCOLORBKGRND_FALSE});
	gl_OC_bUseDefaultColorBkGround[iOCInstance] := bUseDefault;
end;



//
// OpenCandySetCustomBrushColor
// ----------------------------
// Calling this procedure after OpenCandyInit in the Inno Setup
// InitializeWizard callback procedure tells the client to draw
// the OpenCandy loading and offer screens on the specified
// solid background color. The OpenCandy loading screen takes the
// same configuration as the first enabled, non-remnant OpenCandy
// offer screen.
//
// Parameters:
//
//   iOCInstance : Client instance ID
//   szColor     : The solid background color to draw on, in '#RGB' form, where R, G, B
//                 are each zero-padded hex values in range 00-FF
//
// Usage:
//
//   // Draw OpenCandy offer screen one on solid red background color
//   OpenCandySetCustomBrushColor(1, '#FF0000');
//

procedure OpenCandySetCustomBrushColor(iOCInstance:Integer; szColor:OCTString);
begin
	if not _OCValidInstance(iOCInstance) then
		Exit;

	if _OCEnabledAndReady(iOCInstance) then
		#ifdef UNICODE
		_OCDLL_OCSetCustomBrushColorW(gl_OC_szAPIGuid[iOCInstance], szColor);
		#else
		_OCDLL_OCSetCustomBrushColorA(gl_OC_szAPIGuid[iOCInstance], szColor);
		#endif
		gl_OC_szCustomBrushColor[iOCInstance] := szColor;
end;



//
// OpenCandyCustomImagePath
// ------------------------
// Calling this procedure after OpenCandyInit in the Inno Setup
// InitializeWizard callback procedure tells the client to load a
// background image from the specified file, which should be a
// fully-qualified path, and composite the OpenCandy loading and
// or offer screen upon it. The OpenCandy loading screen takes the
// same configuration as the first enabled, non-remnant OpenCandy
// offer screen.  The image dimensions should match the dimensions
// of the main installer window.
//
// Parameters:
//
//   iOCInstance : Client instance ID
//   szImagePath : A fully-qualified path to the background image.
//
// Usage:
//
//   // Use a custom image background for OpenCandy offer screen 1
//   OpenCandyCustomImagePath(1, ExpandConstant('{tmp}\MyInstallerBackground.png'));
//

procedure OpenCandyCustomImagePath(iOCInstance:Integer; szImagePath:OCTString);
begin
	if gl_OC_bNoCandy or not _OCValidInstance(iOCInstance) then
		Exit;
	if gl_OC_bHasBeenInitialized[iOCInstance] then
		#ifdef UNICODE
		_OCDLL_OCSetBkGrdImagePathW(gl_OC_szAPIGuid[iOCInstance], szImagePath);
		#else
		_OCDLL_OCSetBkGrdImagePathA(gl_OC_szAPIGuid[iOCInstance], szImagePath);
		#endif
	gl_OC_szCustomImagePath[iOCInstance] := szImagePath;
end;



//
// _OpenCandyInitInternal
// ----------------------
// This procedure is internal to this helper script. Do not
// call it from your own code. Instead see OpenCandyInit.
//

procedure _OpenCandyInitInternal();
var
	i:Integer;
	iOCWaitOnInstance:Integer;
	iMaxInitTimeMs:Integer;
	iOCInstance:Integer;
begin

	if gl_OC_bNoCandy or gl_OC_bHasDone_Init then
		Exit;

	gl_OC_bHasDone_Init := true;

	// Determine which instance will have any blocking time
	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if ((gl_OC_szAPIGuid[iOCInstance] <> '') and ({#OC_INIT_MODE_NORMAL} = gl_OC_init_iInitModeOffer[iOCInstance])) then
			iOCWaitOnInstance := iOCInstance;


	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if (gl_OC_szAPIGuid[iOCInstance] <> '') and not gl_OC_bNoCandy then
		begin
			// Handle any OpenCandy command line options
			for i := 0 to ParamCount() do
				#ifdef UNICODE
				_OCDLL_OCSetCmdLineValuesW(gl_OC_szAPIGuid[iOCInstance], ParamStr(i));
				#else
				_OCDLL_OCSetCmdLineValues(gl_OC_szAPIGuid[iOCInstance], ParamStr(i));
				#endif


			// Check if command line options have disabled OpenCandy
			if {#OC_CANDY_DISABLED} = _OCDLL_OCGetNoCandy(gl_OC_szAPIGuid[iOCInstance]) then
				_OCSetNoCandy();

			if not gl_OC_bNoCandy then
			begin
				// Pass advanced options to client
				#ifdef OC_ADV_OPTIONS
				#ifdef UNICODE
				_OCDLL_SetClientAdvancedOptionsW(gl_OC_szAPIGuid[iOCInstance], '{#OC_ADV_OPTIONS}');
				#else
				_OCDLL_SetClientAdvancedOptions(gl_OC_szAPIGuid[iOCInstance], '{#OC_ADV_OPTIONS}');
				#endif
				#endif

				iMaxInitTimeMs := 0;
				if iOCInstance = iOCWaitOnInstance then
					if {#OC_INIT_PERFORM_NOW} = gl_OC_init_iPerformInit then
						iMaxInitTimeMs := {#OC_MAX_INIT_TIME_INIT_PERFORM_NOW}
					else
						iMaxInitTimeMs := {#OC_MAX_INIT_TIME_INIT_PERFORM_BYPAGEORDER};


				// Initialize OpenCandy client
				#ifdef UNICODE
				if {#OC_INIT_SUCCESS} = _OCDLL_OCInitW(gl_OC_szAPIGuid[iOCInstance], gl_OC_init_tszProductName, gl_OC_init_tszProductKey[iOCInstance], gl_OC_init_tszProductSecret[iOCInstance],
													   gl_OC_init_tszInstallerLanguage, true, iMaxInitTimeMs, gl_OC_init_iInitModeOffer[iOCInstance], {#OC_INIT_PROGRESSBAR_ENABLE}, {#OC_INIT_PROGRESSBAR_DELAY}) then
				#else
				if {#OC_INIT_SUCCESS} = _OCDLL_OCInit(gl_OC_szAPIGuid[iOCInstance], gl_OC_init_tszProductName, gl_OC_init_tszProductKey[iOCInstance], gl_OC_init_tszProductSecret[iOCInstance],
													  gl_OC_init_tszInstallerLanguage, true, iMaxInitTimeMs, gl_OC_init_iInitModeOffer[iOCInstance], {#OC_INIT_PROGRESSBAR_ENABLE}, {#OC_INIT_PROGRESSBAR_DELAY}) then
				#endif
				begin
					gl_OC_bHasBeenInitialized[iOCInstance] := true;
					// Apply non-default deferred settings
					if  gl_OC_init_iPerformInit <> {#OC_INIT_PERFORM_NOW} then
					begin
						// Inform the client if offers have already been disabled
						if not gl_OC_bOfferIsEnabled[iOCInstance] then
							SetOCOfferEnabled(iOCInstance, gl_OC_bOfferIsEnabled[iOCInstance]);

						// Apply background color settings for loading screen and offer screen
						if not gl_OC_bUseDefaultColorBkGround[iOCInstance] then
							OpenCandySetUseDefaultColorBkgrnd(iOCInstance, gl_OC_bUseDefaultColorBkGround[iOCInstance]);
						if gl_OC_szCustomBrushColor[iOCInstance] <> '' then
							OpenCandySetCustomBrushColor(iOCInstance, gl_OC_szCustomBrushColor[iOCInstance]);
						if gl_OC_szCustomImagePath[iOCInstance] <> '' then
							OpenCandyCustomImagePath(iOCInstance, gl_OC_szCustomImagePath[iOCInstance]);
					end;
				end
				else
					_OCSetNoCandy();
			end;
		end;
end;



// OpenCandyInit
// -------------
// Performs initialization of the OpenCandy DLL
// and checks for offers to present.  This procedure
// must be called before any other.
//
// Parameters:
//
//   tszProductName    : Your publisher or product name (will be provided by OpenCandy)
//   tszKey1           : Your product key for offer 1(will be provided by OpenCandy)
//   tszSecret1        : Your product secret for offer 1(will be provided by OpenCandy)
//   iInitModeOffer1   : The operating mode for OpenCandy offer one. Pass OC_INIT_MODE_NORMAL
//                       for normal operation or OC_INIT_MODE_REMNANT if OpenCandy should not
//                       show offers. Do not use InitMode to handle /NOCANDY or silent installations,
//                       this is done automatically for you.
//   tszKey2           : Your product key for offer 2 (will be provided by OpenCandy)
//   tszSecret2        : Your product secret for offer 2(will be provided by OpenCandy)
//   iInitModeOffer2   : The operating mode for OpenCandy offer two. Pass OC_INIT_MODE_NORMAL
//                       for normal operation or OC_INIT_MODE_REMNANT if OpenCandy should not
//                       show offers. Do not use InitMode to handle /NOCANDY or silent installations,
//                       this is done automatically for you.
//   iPerformInit      : When to perform initialization. Pass OC_INIT_PERFORM_NOW to extract and load
//                       the OpenCandy Network Client library and connect to the OpenCandy network
//                       immediately, or OC_INIT_PERFORM_BYPAGEORDER to defer these operations until
//                       the end user navigates to the OpenCandyLoadDLLPage and OpenCandyConnectPage
//                       placeholder pages, respectively.
//
// Usage (Using sample values for internal testing purposes only):
//
//   procedure InitializeWizard;
//   var
//     OCtszInstallerLanguage: OCTString;
//   begin
//     // Translate language from the value of the "Name"
//     // parameter assigned in the "[Languages]" section
//     // into ISO 639-1 Alpha-2 codes for the OpenCandy API
//     OCtszInstallerLanguage := ActiveLanguage();
//     if(OCtszInstallerLanguage = 'default') then
//       OCtszInstallerLanguage := 'en';
//
//	   // Initialize OpenCandy
//	   OpenCandyInit('OpenCandy Sample',
//	                 '748ad6d80864338c9c03b664839d8161', 'dfb3a60d6bfdb55c50e1ef53249f1198', {#OC_INIT_MODE_NORMAL},
//	                 '748ad6d80864338c9c03b664839d8161', 'dfb3a60d6bfdb55c50e1ef53249f1198', {#OC_INIT_MODE_NORMAL},
//	                 OCtszInstallerLanguage, {#OC_INIT_PERFORM_NOW});
//
//   end;
//

procedure OpenCandyInit(tszProductName, tszKey1, tszSecret1:OCTString; iInitModeOffer1:Integer; tszKey2, tszSecret2:OCTString; iInitModeOffer2:Integer; tszLanguage:OCTString; iPerformInit:Integer);
begin
	// Prepare OpenCandy Inno Setup layer
	_OpenCandyPrepareInnoAPI();

	// Store init options
	gl_OC_init_tszProductName := tszProductName;
	gl_OC_init_tszProductKey[1] := tszKey1;
	gl_OC_init_tszProductSecret[1] := tszSecret1;
	gl_OC_init_tszProductKey[2] := tszKey2;
	gl_OC_init_tszProductSecret[2] := tszSecret2;
	gl_OC_init_tszInstallerLanguage := tszLanguage;
	gl_OC_init_iInitModeOffer[1] := iInitModeOffer1;
	gl_OC_init_iInitModeOffer[2] := iInitModeOffer2;
	gl_OC_init_iPerformInit := iPerformInit;

	// Automatically disable OpenCandy under various circumstances
	_OpenCandyAutoSelfDisable();
	if {#OC_INIT_PERFORM_NOW} = iPerformInit then
	begin
		// Load and initialize OpenCandy immediately
		_OpenCandyLoadDLL();
		_OpenCandyInitInternal();
	end;
end;



// OpenCandyInsertLoadDLLPage
// --------------------------
// Inserts a placeholder page that is used to load the
// OpenCandy Network Client library after the InitializeWizard
// callback procedure when the OC_INIT_PERFORM_BYPAGEORDER option is used.
//
// The placeholder page should be inserted as early as possible in
// the wizard sequence in order to maximize the likelihood
// that offers will be ready by the time the end user reaches the
// offer screen.
//
// The function returns the ID of the new placeholder page that
// is inserted, or the same page ID that is passed to the function as
// iAfterPageID if the new placeholder page is not created successfully.
// The page ID returned can be used to insert other custom pages
// after this placeholder page.
//
// Parameters:
//
//   iAfterPageID : Insert after this page ID
//
// Usage:
//
//   iOpenCandyNewPageID := OpenCandyInsertLoadDLLPage(wpLicense);
//

function OpenCandyInsertLoadDLLPage(iAfterPageID:Integer):Integer;
begin
	Result := iAfterPageID;
	if not (gl_OC_bNoCandy or (gl_OC_objOCLoadDLLPage <> nil) or (gl_OC_init_iPerformInit <> {#OC_INIT_PERFORM_BYPAGEORDER})) then
		gl_OC_objOCLoadDLLPage := CreateCustomPage(iAfterPageID, ' ', ' ');
	if gl_OC_objOCLoadDLLPage <> nil then
		Result := gl_OC_objOCLoadDLLPage.ID;
end;



// OpenCandyInsertConnectPage
// --------------------------
// Inserts a placeholder page that is used to connect the OpenCandy
// Network Client to the OpenCandy Network after the InitializeWizard
// callback when the OC_INIT_PERFORM_BYPAGEORDER option is used.
//
// This placeholder page must be inserted after the placeholder page
// that was inserted by the OpenCandyInsertLoadDLLPage function, and in
// most circumstances should follow immediately afterwards. Otherwise,
// it should be inserted as early as possible in the wizard sequence
// in order to maximize the likelihood that offers will be ready by the
// time the end user reaches the offer screen.
//
// The function returns the ID of the new placeholder page that
// is inserted, or the same page ID that is passed to the function as
// iAfterPageID if the new placeholder page is not created successfully.
// The page ID returned can be used to insert other custom pages
// after this placeholder page.
//
//
// Parameters:
//
//   iAfterPageID : Insert after this page ID
//
// Usage:
//
//   iOpenCandyNewPageID := OpenCandyInsertLoadDLLPage(wpLicense);
//   iOpenCandyNewPageID := OpenCandyInsertConnectPage(iOpenCandyNewPageID);
//

function OpenCandyInsertConnectPage(iAfterPageID:Integer):Integer;
begin
	Result := iAfterPageID;
	if not (gl_OC_bNoCandy or (gl_OC_objOCConnectPage <> nil) or (gl_OC_init_iPerformInit <> {#OC_INIT_PERFORM_BYPAGEORDER})) then
		gl_OC_objOCConnectPage := CreateCustomPage(iAfterPageID, ' ', ' ');
	if gl_OC_objOCConnectPage <> nil then
		Result := gl_OC_objOCConnectPage.ID;
end;



//
// OpenCandyInsertLoadingPage
// --------------------------
// Inserts a placeholder page that is used to display a loading
// screen while the OpenCandy client is retrieving offers from
// the OpenCandy network. This placeholder page must be inserted after
// the page inserted by OpenCandyInsertConnectPage.
//
// The placeholder page should generally be inserted immediately before
// the OpenCandy offer screen to minimize both the likelihood that it
// will be displayed and the display duration.
//
// The function returns the ID of the new placeholder page that
// is inserted, or the same page ID that is passed to the function as
// iAfterPageID if the new placeholder page is not created successfully.
// The page ID returned can be used to insert other custom pages
// after this placeholder page.
//
//
// Parameters:
//
//   iAfterPageID  : Insert after this page ID
//   szCaption     : The page caption, or ' ' to display no caption
//   szDescription : The page description, or ' ' to display no description
//   szMessage     : The message displayed on the loading screen, or ' ' to display no message
//   szFontFace    : The font face used to display the message on the loading screen
//   iFontSize     : The font size used to display the message on the loading screen
//
// Usage:
//
//   iOpenCandyNewPageID := OpenCandyInsertConnectPage(wpLicense);
//   iOpenCandyNewPageID := OpenCandyInsertLoadDLLPage(iOpenCandyNewPageID);
//   // ...
//   iOpenCandyNewPageID := OpenCandyInsertLoadingPage(wpSelectTasks, ' ', ' ', 'Loading...', 'Arial', 100);
//

function OpenCandyInsertLoadingPage(iAfterPageID:Integer; szCaption, szDescription, szMessage, szFontFace:OCTString; iFontSize:Integer):Integer;
begin
	Result := iAfterPageID;
	if not (gl_OC_bNoCandy or (gl_OC_objLoadingPage <> nil)) then
	begin
		gl_OC_objLoadingPage := CreateCustomPage(iAfterPageID, szCaption, szDescription);
		gl_OC_szLoadingMsg := szMessage;
		gl_OC_szFontName := szFontFace;
		gl_OC_iFontSize := iFontSize;
	end;
	if gl_OC_objLoadingPage <> nil then
		Result := gl_OC_objLoadingPage.ID;
end;



// OpenCandyInsertOfferPage / OpenCandyInsertOfferPage2
// ----------------------------------------------------
// Inserts the OpenCandy offer pages. The offer pages are displayed if
// offers have become ready in the time since the placeholder page
// inserted using OpenCandyInsertConnectPage was reached and a valid
// offer was found for the specific end user system.
//
// The offer pages should generally be inserted as late in the installation
// sequence as possible, just before installation begins. This helps to
// maximize the likelihood that offers will be ready.
//
// The function returns the ID of the new placeholder page that
// is inserted, or the same page ID that is passed to the function as
// iAfterPageID if the new placeholder page is not created successfully.
// The page ID returned can be used to insert other custom pages
// after the offer page.
//
//
// Parameters:
//
//   iOCInstance : Client instance ID
//   iAfterPageID : Insert after this page ID
//
// Usage:
//
//   iOpenCandyNewPageID := OpenCandyInsertConnectPage(wpLicense);
//   iOpenCandyNewPageID := OpenCandyInsertLoadDLLPage(iOpenCandyNewPageID);
//   // ...
//   iOpenCandyNewPageID := OpenCandyInsertLoadingPage(wpSelectTasks, ' ', ' ', 'Loading...', 'Arial', 100);
//   iOpenCandyNewPageID := OpenCandyInsertOfferPage(iOpenCandyNewPageID);
//

function OpenCandyInsertOfferPage_Instance(iOCInstance, iAfterPageID:Integer):Integer;
begin
	Result := iAfterPageID;
	if not _OCValidInstance(iOCInstance) then
		Exit;
	if not (gl_OC_bNoCandy or (gl_OC_objOCOfferPage[iOCInstance] <> nil)) then
		gl_OC_objOCOfferPage[iOCInstance] := CreateCustomPage(iAfterPageID, ' ', ' ');
	if gl_OC_objOCOfferPage[iOCInstance] <> nil then
		Result := gl_OC_objOCOfferPage[iOCInstance].ID;
end;

function OpenCandyInsertOfferPage(iAfterPageID:Integer):Integer;
begin
	Result := OpenCandyInsertOfferPage_Instance(1, iAfterPageID);
end;

function OpenCandyInsertOfferPage2(iAfterPageID:Integer):Integer;
begin
	Result := OpenCandyInsertOfferPage_Instance(2, iAfterPageID);
end;



//
// GetOCOfferStatus
// ----------------
// Allows you to determine if an offer is currently available. This is
// done automatically for you before the offer screen is shown. Typically
// you do not need to call this function from your own code directly.
//
// The offer status is placed on the stack and may be one of:
// {#OC_OFFER_STATUS_CANOFFER_READY}    - An OpenCandy offer is available and ready to be shown
// {#OC_OFFER_STATUS_CANOFFER_NOTREADY} - An offer is available but is not yet ready to be shown
// {#OC_OFFER_STATUS_QUERYING_NOTREADY} - The remote API is still being queried for offers
// {#OC_OFFER_STATUS_NOOFFERSAVAILABLE} - No offers are available
//
// When calling this function you must indicate whether the information returned
// will be used to decide whether the OpenCandy offer screen will be shown, e.g.
// if the information may result in a call to SetOCOfferEnabled. This helps
// to optimize future OpenCandy SDKs for better performance with your product.
//
// Parameters:
//
//   iOCInstance             : Client instance ID
//   bDeterminesOfferEnabled : Will the result of this specific call determine whether the
//                             offer screen will be shown?
//
// Usage:
//
//   // Test if OpenCandy offer 1 is ready to be shown.
//   // Indicate the result is informative only and the result of this call
//   // is not actually used to decide whether the offer will be shown or not.
//   if {#OC_OFFER_STATUS_CANOFFER_READY} = GetOCOfferStatus(1, false) then
//

function GetOCOfferStatus(iOCInstance:Integer; bDeterminesOfferEnabled:Boolean):Integer;
begin
	Result := {#OC_OFFER_STATUS_NOOFFERSAVAILABLE};
	if not _OCValidInstance(iOCInstance) then
		Exit;

	if not _OCEnabledAndReady(iOCInstance) then
		Exit;

	Result := _OCDLL_OCGetAsyncOfferStatus(gl_OC_szAPIGuid[iOCInstance], bDeterminesOfferEnabled);
end;



//
// _OpenCandyUpdateCaptions
// ------------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// Update the Title and Description text on the OpenCandy offer screen
// using strings from the server (preferred) or else fall-back strings.
// The result is true if strings from the server were ready.
//
// Parameters:
//
//   iOCInstance          : Client instance ID
//   objTWizardPage       : The OpenCandy wizard page
//   szDefaultCaption     : The default page caption, if no banner strings from the server are ready
//   szDefaultDescription : The default page description, if no banner strings from the server are ready
//
// Usage:
//
//   _OpenCandyUpdateCaptions(gl_OC_objOCOfferPage[1], ' ', ' ');
//

function _OpenCandyUpdateCaptions(iOCInstance:Integer; var objTWizardPage:TWizardPage; szDefaultCaption, szDefaultDescription:OCTString):Boolean;
var
	iBannerInfoResult:Integer;
	tszDesc: OCTString;
	tszTitle: OCTString;
begin
	iBannerInfoResult := {#OC_OFFER_BANNER_FOUNDNEITHER};
	if _OCEnabledAndReady(iOCInstance) then
	begin
		tszTitle := StringOfChar(' ', {#OC_STR_CHARS});
		tszDesc := StringOfChar(' ', {#OC_STR_CHARS});
		#ifdef UNICODE
			iBannerInfoResult := _OCDLL_OCGetBannerInfoW(gl_OC_szAPIGuid[iOCInstance], gl_OC_szMemMapGuid[iOCInstance], tszTitle, tszDesc);
		#else
			iBannerInfoResult := _OCDLL_OCGetBannerInfo(gl_OC_szAPIGuid[iOCInstance], gl_OC_szMemMapGuid[iOCInstance], tszTitle, tszDesc);
		#endif

		case iBannerInfoResult of
			{#OC_OFFER_BANNER_FOUNDTITLE}: tszDesc := ' ';
			{#OC_OFFER_BANNER_FOUNDDESCRIPTION}: tszTitle := ' ';
			{#OC_OFFER_BANNER_FOUNDNEITHER}:
			begin
				tszTitle := szDefaultCaption;
				tszDesc := szDefaultDescription;
			end;
		end;
		objTWizardPage.Caption := tszTitle;
		objTWizardPage.Description := tszDesc;
	end;
	Result := iBannerInfoResult <> {#OC_OFFER_BANNER_FOUNDNEITHER};
end;



//
// _OpenCandyShowLoadingScreen
// ---------------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// Shows the loading screen while the OpenCandy client waits
// for the offer availability state to become final, or the maximum
// wait time to elapse. The user is prevented from clicking 'Next'
// while the loading screen is displayed.
//

procedure _OpenCandyShowLoadingScreen();
var
	OC_iMaxLoadingTimeMs:Integer;
	iOCInstance: Integer;
begin
	Wizardform.NextButton.Enabled := false;
	Wizardform.CancelButton.Enabled := false;
	gl_OC_bHasDone_LoadingScreen  := true;
	gl_OC_bAutoNextAfterLoadingScreen := true;

	if {#OC_INIT_PERFORM_NOW} = gl_OC_init_iPerformInit then
		OC_iMaxLoadingTimeMs := {#OC_MAX_LOADING_TIME_INIT_PERFORM_NOW}
	else
		OC_iMaxLoadingTimeMs := {#OC_MAX_LOADING_TIME_INIT_PERFORM_BYPAGEORDER};

	gl_OC_iLoadingScreenInstance := 0;
	for iOCInstance := {#OC_MAX_INSTANCES} downto 1 do
		if gl_OC_bHasBeenInitialized[iOCInstance]
		   and ({#OC_INIT_MODE_NORMAL} = gl_OC_init_iInitModeOffer[iOCInstance])
		   and gl_OC_bOfferIsEnabled[iOCInstance] then
			gl_OC_iLoadingScreenInstance := iOCInstance;

	if gl_OC_iLoadingScreenInstance <> 0 then
	begin
		#ifdef UNICODE
		_OCDLL_OCShowLoadingScreen2W(gl_OC_szAPIGuid[gl_OC_iLoadingScreenInstance], OC_iMaxLoadingTimeMs, gl_OC_objLoadingPage.Surface.Handle, gl_OC_szLoadingMsg, gl_OC_szFontName, gl_OC_iFontSize);
		#else
		_OCDLL_OCShowLoadingScreen2A(gl_OC_szAPIGuid[gl_OC_iLoadingScreenInstance], OC_iMaxLoadingTimeMs, gl_OC_objLoadingPage.Surface.Handle, gl_OC_szLoadingMsg, gl_OC_szFontName, gl_OC_iFontSize);
		#endif
		gl_OC_iLoadingScreenInstance := 0;
	end;

	Wizardform.NextButton.Enabled := true;
	Wizardform.CancelButton.Enabled := true;
	if gl_OC_bAutoNextAfterLoadingScreen then
		Wizardform.NextButton.OnClick(nil);
end;



//
// _OpenCandyUpdateExternalSkipState
// ---------------------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// If necessary, queries the publisher installer script via the SkipQuery
// callback to determine if a Skip control was shown prior to the OpenCandy
// offer screens in the wizard sequence, and if that control was activated.
//

procedure _OpenCandyUpdateExternalSkipState(bIsOfferPage: Boolean);
begin
	// Query external skip state
	if (gl_OC_fnCallbackSkipQuery <> nil) and not
	   (({#OC_SKIPSHOWN_NOTSHOWN_EXTERNAL} = gl_OC_iSkipShown) or
	    ({#OC_SKIPSHOWN_SHOWN_OPENCANDY} = gl_OC_iSkipShown) or
	    ({#OC_SKIPSTATE_SKIPPED_EXTERNAL} = gl_OC_iSkipState)) then
		Case gl_OC_fnCallbackSkipQuery() of
			{#OC_SKIPQUERY_SHOWN_ACTIVATED}:
			begin
				gl_OC_iSkipShown := {#OC_SKIPSHOWN_SHOWN_EXTERNAL}
				gl_OC_iSkipState := {#OC_SKIPSTATE_SKIPPED_EXTERNAL}
			end;
			{#OC_SKIPQUERY_SHOWN_NOTACTIVATED}:
				gl_OC_iSkipShown := {#OC_SKIPSHOWN_SHOWN_EXTERNAL};
			{#OC_SKIPQUERY_NOTSHOWN}:
				if bIsOfferPage and ({#OC_SKIPSHOWN_UNKNOWN} = gl_OC_iSkipShown) then
					gl_OC_iSkipShown := {#OC_SKIPSHOWN_NOTSHOWN_EXTERNAL};
		end;
end;



//
// _OpenCandyHasShownSkipControl
// -----------------------------
// This function is internal to this helper script. Do not
// call it from your own code.
//
// If necessary, notify the publisher installer script that a Skip control was
// shown by OpenCandy, and whether or not it was activated by the user.
//

procedure _OpenCandyHasShownSkipControl(iSkipNotification: Integer);
begin
	// Inform publisher installer of skip control
	if (gl_OC_fnCallbackSkipNotify <> nil) then
		if ({#OC_SKIPNOTIFICATION_SHOWN_ACTIVATED} = iSkipNotification) then
			gl_OC_fnCallbackSkipNotify({#OC_SKIPNOTIFICATION_SHOWN_ACTIVATED})
		else if ({#OC_SKIPSHOWN_SHOWN_OPENCANDY} <>  gl_OC_iSkipShown) then
			gl_OC_fnCallbackSkipNotify({#OC_SKIPNOTIFICATION_SHOWN_NOTACTIVATED});
	gl_OC_iSkipShown := {#OC_SKIPSHOWN_SHOWN_OPENCANDY};
end;



//
// OpenCandyShouldSkipPage()
// -------------------------
//
// This function needs to be called from the ShouldSkipPage Inno script
// event function so that Inno Setup can determine whether the OpenCandy
// offer page should be displayed. The function returns true if the
// current page is the OpenCandy offer page and no offer is to be
// presented.
//
// Usage:
//
//   function ShouldSkipPage(PageID:Integer):Boolean;
//   begin
//     Result := false; // Don't skip pages by default
//
//     if OpenCandyShouldSkipPage(PageID) then
//       Result := true;
//   end;
//

function OpenCandyShouldSkipPage(PageID:Integer):Boolean;
var
	iOCInstance:Integer;
begin
	Result := false;

	// Handle deferred loading of the OpenCandy Network Client DLL
	if ({#OC_INIT_PERFORM_BYPAGEORDER} = gl_OC_init_iPerformInit) and (gl_OC_objOCLoadDLLPage <> nil) then
		if (PageID = gl_OC_objOCLoadDLLPage.ID) then
		begin
			_OpenCandyLoadDLL();
			Result := true;
		end;

	// Handle deferred connection to the OpenCandy Network
	if ({#OC_INIT_PERFORM_BYPAGEORDER} = gl_OC_init_iPerformInit) and (gl_OC_objOCConnectPage <> nil) then
		if PageID = gl_OC_objOCConnectPage.ID then
		begin
			_OpenCandyInitInternal();
			Result := true;
		end;

	// Handle reaching the OpenCandy loading page
	if  gl_OC_objLoadingPage <> nil then
		if PageID = gl_OC_objLoadingPage.ID then
		begin
			Result := true;
			if (not gl_OC_bHasDone_LoadingScreen) and _OCEnabledAndReady(0) and ({#OC_SKIPSTATE_NOTSKIPPED} = gl_OC_iSkipState) then
			begin
				_OpenCandyUpdateExternalSkipState(false);
				Result := gl_OC_iSkipState <> {#OC_SKIPSTATE_NOTSKIPPED};
			end;
		end;

	// Handle reaching the OpenCandy offer pages
	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if gl_OC_objOCOfferPage[iOCInstance] <> nil then
			if PageID = gl_OC_objOCOfferPage[iOCInstance].ID then
			begin
				if _OCEnabledAndReady(iOCInstance) then
				begin
					// Test if an offer is ready
					if not gl_OC_bHasReachedOCPage[iOCInstance] then
						gl_OC_bUseOfferPage[iOCInstance] := {#OC_OFFER_STATUS_CANOFFER_READY} = _OCDLL_OCGetAsyncOfferStatus(gl_OC_szAPIGuid[iOCInstance], true);
				end
				else
					gl_OC_bUseOfferPage[iOCInstance] := false;
				gl_OC_bHasReachedOCPage[iOCInstance] := true;

				// Test if this page has already been skipped
				if gl_OC_bUseOfferPage[iOCInstance] then
				begin
					// Update external skip state
					_OpenCandyUpdateExternalSkipState(true);
					if {#OC_SKIPSTATE_NOTSKIPPED} <> gl_OC_iSkipState then
					begin
						gl_OC_bUseOfferPage[iOCInstance] := false;
						if {#OC_SKIPSTATE_SKIPPED_OPENCANDY} = gl_OC_iSkipState then
							_OCDLL_OCSkipOffer(gl_OC_szAPIGuid[iOCInstance], {#OC_SKIPPED_OCSKIPNONOWNERSCREEN})
						else if {#OC_SKIPSTATE_SKIPPED_EXTERNAL} = gl_OC_iSkipState then
							_OCDLL_OCSkipOffer(gl_OC_szAPIGuid[iOCInstance], {#OC_SKIPPED_3RDPARTYSKIP});
					end;
				end;
				Result := not gl_OC_bUseOfferPage[iOCInstance];
			end;
end;



//
// OpenCandyCurPageChanged
// -----------------------
// This function needs to be called from CurPageChanged() Inno script
// event function so that the OpenCandy pages work correctly.
//
// Usage:
//
//   procedure CurPageChanged(CurPageID:Integer);
//   begin
//     OpenCandyCurPageChanged(CurPageID);
//   end;
//

procedure OpenCandyCurPageChanged(CurPageID:Integer);
var
	iOCInstance:Integer;
	i: Integer;
begin
	if not _OCEnabledAndReady(0) then
		Exit;

	// Handle showing the OpenCandy loading screen
	if (gl_OC_objLoadingPage <> nil) then
		if (CurPageID = gl_OC_objLoadingPage.ID) then
		begin
			_OpenCandyShowLoadingScreen();
			Exit;
		end;

	// Handle detaching the OpenCandy offer screens
	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if (gl_OC_objOCOfferPage[iOCInstance] <> nil) then
			if (CurPageID <> gl_OC_objOCOfferPage[iOCInstance].ID) and gl_OC_bAttached[iOCInstance] then
				if _OCEnabledAndReady(iOCInstance) then
				begin
					_OCDLL_OCDetach(gl_OC_szAPIGuid[iOCInstance]);
					gl_OC_bAttached[iOCInstance] := false;
				end;

	// Handle showing the OpenCandy offer screens
	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if (gl_OC_objOCOfferPage[iOCInstance] <> nil) and gl_OC_bUseOfferPage[iOCInstance] then
			if (CurPageID = gl_OC_objOCOfferPage[iOCInstance].ID) and not gl_OC_bAttached[iOCInstance] then
				if _OCEnabledAndReady(iOCInstance) then
				begin
					// Check if OpenCandy should display skip controls
					if (gl_OC_fnCallbackSkipEnable <> nil) then
					begin
						if (({#OC_SKIPSHOWN_NOTSHOWN_EXTERNAL} = gl_OC_iSkipShown)
							or ((gl_OC_fnCallbackSkipQuery = nil) and ({#OC_SKIPSHOWN_UNKNOWN} = gl_OC_iSkipShown))
							) and ({#OC_SKIPDISPLAY_INSTANCE_NOTSET} = gl_OC_iSkipDisplayInstance) then
						begin
							Case gl_OC_fnCallbackSkipEnable() of
							{#OC_SKIP_ENABLED_ALWAYS}:
								gl_OC_iSkipDisplayInstance := iOCInstance;
							{#OC_SKIP_ENABLED_OCMULTIOFFERONLY}:
								begin
									for i := iOCInstance + 1 to {#OC_MAX_INSTANCES} do
										if {#OC_OFFER_STATUS_NOOFFERSAVAILABLE} <> GetOCOfferStatus(i, false) then
											gl_OC_iSkipDisplayInstance := iOCInstance;
								end;
							else
								gl_OC_iSkipDisplayInstance := {#OC_SKIPDISPLAY_INSTANCE_DISABLEALL};
							end;
						end;
					end;

					// Set the page caption and description
					_OpenCandyUpdateCaptions(iOCInstance, gl_OC_objOCOfferPage[iOCInstance], ' ', ' ');

					// Adjust the window
					_OCDLL_OCAdjustPage(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle, {#OC_WND_OFFSET_X}, {#OC_WND_OFFSET_Y}, {#OC_WND_WIDTH}, {#OC_WND_HEIGHT});

					// Prepare Decline control
					_OCDLL_OCSetDeclineButton(gl_OC_szAPIGuid[iOCInstance], WizardForm.BackButton.Handle, {#OC_CTRL_DECLINE_ALIGNMENT}, {#OC_CTRL_DECLINE_OFFSET_X}, {#OC_CTRL_DECLINE_OFFSET_Y}, WizardForm.NextButton.Handle, '{#OC_CTRL_DECLINE_FONT_NAME}', {#OC_CTRL_DECLINE_FONT_SIZE});

					// Show Skip control if necessary
					if (iOCInstance = gl_OC_iSkipDisplayInstance) then
					begin
						#ifdef UNICODE
						_OCDLL_OCShowSkipAllButtonW(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle,
													{#OC_CTRL_SKIP_ALIGNMENT}, {#OC_CTRL_SKIP_OFFSET_X}, {#OC_CTRL_SKIP_OFFSET_Y},
													WizardForm.NextButton.Handle, {#OC_CTRL_SKIP_LINE_ENABLE}, '{#OC_CTRL_SKIP_LINE_FGCOLOR}', '{#OC_CTRL_SKIP_FONT_NAME}', {#OC_CTRL_SKIP_FONT_SIZE}, {#OC_CTRL_SKIP_UNDERLINE},
													'{#OC_CTRL_SKIP_TEXT_FGCOLOR}', '{#OC_CTRL_SKIP_TEXT_BGCOLOR}', '{#OC_CTRL_SKIP_TEXT_CUSTOM}');
						#else
						_OCDLL_OCShowSkipAllButtonA(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle,
													{#OC_CTRL_SKIP_ALIGNMENT}, {#OC_CTRL_SKIP_OFFSET_X}, {#OC_CTRL_SKIP_OFFSET_Y},
													WizardForm.NextButton.Handle, {#OC_CTRL_SKIP_LINE_ENABLE}, '{#OC_CTRL_SKIP_LINE_FGCOLOR}', '{#OC_CTRL_SKIP_FONT_NAME}', {#OC_CTRL_SKIP_FONT_SIZE}, {#OC_CTRL_SKIP_UNDERLINE},
													'{#OC_CTRL_SKIP_TEXT_FGCOLOR}', '{#OC_CTRL_SKIP_TEXT_BGCOLOR}', '{#OC_CTRL_SKIP_TEXT_CUSTOM}');
						#endif
					end;

					// Attach the OpenCandy offer screen
					gl_OC_bAttached[iOCInstance] := {#OC_OFFER_RUNDLG_FAILURE} <> _OCDLL_OCRunDialog(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle);
					if not gl_OC_bAttached[iOCInstance] then
					begin
						// Disable all offers because of errors showing offer screens
						_OCDLL_OCRestorePage(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle);
						for i := 1 to {#OC_MAX_INSTANCES} do
						begin
							gl_OC_bUseOfferPage[i] := false;
							SetOCOfferEnabled(i, false);
						end;
						Wizardform.NextButton.OnClick(nil);
					end;
				end;
end;



//
// OpenCandyNextButtonClick
// ------------------------
// This function needs to be called be called from the NextButtonClick()
// Inno script event function so that Inno Setup does not allow an end user
// to proceed past the OpenCandy offer screen in the event that the user
// must make a selection and hasn't yet done so. The function returns false
// if the user should not be allowed to proceed.
//
// Usage:
//
//   function NextButtonClick(CurPageID:Integer):Boolean;
//   begin
//     Result := true; // Allow action by default
//     if not OpenCandyNextButtonClick(CurPageID) then
//       Result := false;
//   end;
//

function OpenCandyNextButtonClick(CurPageID:Integer):Boolean;
var
	iOCInstance: Integer;
	nSkipButtonClicked: Integer;
begin
	Result := true;

	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if (gl_OC_objOCOfferPage[iOCInstance] <> nil) then
			if (CurPageID = gl_OC_objOCOfferPage[iOCInstance].ID) then
				if _OCEnabledAndReady(iOCInstance) then
				begin
					// Test if a decline or skip control was activated
					if (gl_OC_iSkipDisplayInstance = iOCInstance) then
						if ({#OC_CONTROL_SKIP_USED} = _OCDLL_OCWasSkipOffersClicked(gl_OC_szAPIGuid[iOCInstance])) then
						begin
							// Skip offers by Skip control
							_OpenCandyHasShownSkipControl({#OC_SKIPNOTIFICATION_SHOWN_ACTIVATED});
							gl_OC_iSkipState := {#OC_SKIPSTATE_SKIPPED_OPENCANDY};
							gl_OC_bUseOfferPage[iOCInstance] := false;
							_OCDLL_OCRestorePage(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle);
							_OCDLL_OCSkipOffer(gl_OC_szAPIGuid[iOCInstance], {#OC_SKIPPED_OCSKIPOWNERSCREEN});
						end
						else
							_OpenCandyHasShownSkipControl({#OC_SKIPNOTIFICATION_SHOWN_NOTACTIVATED});

					if (gl_OC_iSkipState = {#OC_SKIPSTATE_NOTSKIPPED}) then
					begin
						if ({#OC_CONTROL_DECLINE_USED} = _OCDLL_OCWasDeclineOfferClicked(gl_OC_szAPIGuid[iOCInstance])) then
						begin
							gl_OC_bUseOfferPage[iOCInstance] := false;
							_OCDLL_OCRestorePage(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle);
							_OCDLL_OCSkipOffer(gl_OC_szAPIGuid[iOCInstance], {#OC_SKIPPED_OPENCANDYDECLINE});
						end
						else if {#OC_OFFER_LEAVEPAGE_DISALLOWED} = _OCDLL_OCCanLeaveOfferPage(gl_OC_szAPIGuid[iOCInstance], gl_OC_szMemMapGuid[iOCInstance]) then
							Result := false
						else
							_OCDLL_OCRestorePage(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle);
					end;
			end;
end;



//
// OpenCandyBackButtonClick
// ------------------------
// This function should be called from BackButtonClick() Inno script
// event function. It restores the layout of the installer window after
// an OpenCandy offer page has been displayed.
//
// Usage:
//
//   function BackButtonClick(CurPageID:Integer):Boolean;
//   begin
//     Result := true; // Allow action by default
//     OpenCandyBackButtonClick(CurPageID);
//   end;
//

procedure OpenCandyBackButtonClick(CurPageID:Integer);
var
	iOCInstance: Integer;
begin

	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if (gl_OC_objOCOfferPage[iOCInstance] <> nil) then
			if (CurPageID = gl_OC_objOCOfferPage[iOCInstance].ID) and _OCEnabledAndReady(iOCInstance) then
			begin
				// Update skip shown state
				if (iOCInstance = gl_OC_iSkipDisplayInstance) then
					_OpenCandyHasShownSkipControl({#OC_SKIPNOTIFICATION_SHOWN_NOTACTIVATED});
				_OCDLL_OCRestorePage(gl_OC_szAPIGuid[iOCInstance], gl_OC_objOCOfferPage[iOCInstance].Surface.Handle);
			end;

	if (gl_OC_objLoadingPage <> nil) then
		if CurPageID = gl_OC_objLoadingPage.ID then
		begin
			// Do not skip to next page after loading screen
			gl_OC_bAutoNextAfterLoadingScreen := false;

			// Unblock the loading screen
			if (gl_OC_iLoadingScreenInstance <> 0) then
				if _OCEnabledAndReady(gl_OC_iLoadingScreenInstance) then
					_OCDLL_OCHideLoadingScreen(gl_OC_szAPIGuid[gl_OC_iLoadingScreenInstance]);
		end;
end;



//
// _OpenCandyExecOfferInternal
// ---------------------------
// This procedure is internal to this helper script. Do not
// call it from your own code.
// Parameters:
//
//   iOCInstance : Client instance ID
//

procedure _OpenCandyExecOfferInternal(iOCInstance:Integer);
begin
		_OCDLL_OCPrepareDownload(gl_OC_szAPIGuid[iOCInstance]);
		if _OCDLL_OCGetOfferState(gl_OC_szAPIGuid[iOCInstance]) = {#OC_OFFER_CHOICE_ACCEPTED} then
			_OCDLL_OCStartDLMgr2Download(gl_OC_szAPIGuid[iOCInstance]);
end;



//
// OpenCandyCurStepChanged
// -----------------------
// This should be called from CurStepChanged() Inno script event function.
// It handles necessary operations at the various different stages of the setup,
// such as installing any offer the user may have accepted.
//
// Usage:
//
//   procedure CurStepChanged(CurStep:TSetupStep);
//   begin
//     OpenCandyCurStepChanged(CurStep);
//   end;
//

procedure OpenCandyCurStepChanged(CurStep:TSetupStep);
var
	iOCInstance: Integer;
begin
	if gl_OC_bNoCandy then
		Exit;

	// ssInstall is just before the product installation starts
	if (CurStep = ssInstall) then
		for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
			if gl_OC_bUseOfferPage[iOCInstance] then
			begin
				if {#OC_OFFER_TYPE_EMBEDDED} = _OCDLL_OCGetOfferType(gl_OC_szAPIGuid[iOCInstance]) then
					_OpenCandyExecOfferInternal(iOCInstance);
					gl_OC_bHasDone_ExecOfferEmbedded[iOCInstance] := true;
			end;


	// ssDone is just before Setup terminates after a successful install
	if (CurStep = ssDone) and not gl_OC_bHasDone_ssDone then
	begin
		gl_OC_bHasDone_ssDone := true;
		gl_OC_bProductInstallSuccess := true;
		for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		begin
			if gl_OC_bUseOfferPage[iOCInstance] then
			begin
				if {#OC_OFFER_TYPE_NORMAL} <> _OCDLL_OCGetOfferType(gl_OC_szAPIGuid[iOCInstance]) then
					if not gl_OC_bHasDone_ExecOfferEmbedded[iOCInstance] then
						_OpenCandyDevModeMsg_Instance(iOCInstance, 'Information:Using embedded mode fallback.', false, 0);
				_OpenCandyExecOfferInternal(iOCInstance);
			end;
			if gl_OC_bHasBeenInitialized[iOCInstance] then
				_OCDLL_OCSignalProductInstalled(gl_OC_szAPIGuid[iOCInstance]);
		end;
	end;
end;



//
// OpenCandyDeinitializeSetup
// --------------------------
// This should be called from DeinitializeSetup() Inno script event function.
// It signals product installation success or failure, and cleans up the
// OpenCandy library.
//
// Usage:
//   procedure DeinitializeSetup();
//   begin
//     OpenCandyDeinitializeSetup();
//   end;
//

procedure OpenCandyDeinitializeSetup();
var
	iOCInstance: Integer;
begin
	if not gl_OC_bNoCandy then
	begin
		if gl_OC_iLoadingScreenInstance <> 0 then
		begin
			gl_OC_bAutoNextAfterLoadingScreen := false;
			if _OCEnabledAndReady(gl_OC_iLoadingScreenInstance) then
					_OCDLL_OCHideLoadingScreen(gl_OC_szAPIGuid[gl_OC_iLoadingScreenInstance]);
		end;

		for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
			if gl_OC_bHasBeenInitialized[iOCInstance] then
			begin
				if not gl_OC_bProductInstallSuccess then
					_OCDLL_OCSignalProductFailed(gl_OC_szAPIGuid[iOCInstance]);
				if gl_OC_bAttached[iOCInstance] then
				begin
					_OCDLL_OCDetach(gl_OC_szAPIGuid[iOCInstance]);
					gl_OC_bAttached[iOCInstance] := false;
				end;
			end;
	end;

	for iOCInstance := 1 to {#OC_MAX_INSTANCES} do
		if (gl_OC_szAPIGuid[iOCInstance] <> '') then
			_OCDLL_OCShutdown(gl_OC_szAPIGuid[iOCInstance]);
end;



//---------------------------------------------------------------------------//
//                    END of OpenCandy Helper Include file                   //
//---------------------------------------------------------------------------//