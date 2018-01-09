VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  ' ==================================================================================================
  ' clsToolTip.cls
  '
  ' TimoSoft ToolTip class
  ' written by Timo "TimoSoft" Kunze (Timo@TimoSoft-Software.com)
  '
  ' This is an implementation of the Windows ToolTip control including Unicode support.
  '
  ' release history:
  ' ----------------
  '  2.1 (03/19/2011) Fixes for DEP compatibility. Thanks to Arndt Mühlenfeld.
  '                   Applied a patch from Gerstgrasser which prevents flickering on Windows 7.
  '                   Code cleanup.
  '  2.0 (02/15/2008) Rewrote large parts of the code. Now Windows Vista is fully supported and the
  '                   code is much cleaner.
  '                   Removed Unicode property. Use the Unicode preprocessor variable instead.
  '                   Added the UseVisualStyle property. It will work on Vista only.
  '                   Added some new tooltip title icons. They will work on Vista only.
  '  1.4 (10/03/2006) Delay times couldn't be changed after tooltip creation.
  '  1.3 (08/20/2004) Added hBitmap support.
  '                   Bugfixes
  '  1.2 (03/21/2004) Removed unused declarations.
  '  1.1 (02/03/2004) Added Attach() and Detach() methods.
  '                   Property Get Unicode() uses IsWindowUnicode() now.
  '                   Some small code cleanups.
  '  1.0 (01/29/2004) Initial release.
  '
  ' known issues:
  ' -------------
  '  - with comctl32.dll 6.0 the tooltip won't reappear under certain circumstances
  '  - hBitmap support doesn't work with all themes, e. g. Vista's Aero theme
  '  - hBitmap support doesn't work with BalloonStyle = True
  '  - hBitmap support doesn't draw the close button if ShowCloseButton = True
  '  - NeedText(): StoreData doesn't have any effect.
  '
  ' --------------------------------------------------------------------------------------------------
  ' visit http://www.TimoSoft-Software.de
  ' ==================================================================================================

  #Const CustomDraw = True
  #Const FullFeatured = True
  #Const Unicode = True

  #Const NeedsSubclassing = CustomDraw Or FullFeatured


  #If CustomDraw Then
    Public Enum BitmapPositionConstants
      bpLeft
      bpTop
      bpRight
      bpBottom
    End Enum
  #End If

  #If FullFeatured Then
    Public Enum ToolTipTitleIconConstants
      tttiNone = 0  ' TTI_NONE
      tttiInfo = 1  ' TTI_INFO
      tttiWarning = 2  ' TTI_WARNING
      tttiError = 3  ' TTI_ERROR
      ' the following require Vista
      tttiInfoLarge = 4  ' TTI_INFO_LARGE
      tttiWarningLarge = 5  ' TTI_WARNING_LARGE
      tttiErrorLarge = 6  ' TTI_ERROR_LARGE
    End Enum
  #End If


  Private Const CLR_INVALID = -1
  Private Const LPSTR_TEXTCALLBACKA = -1
  Private Const LPSTR_TEXTCALLBACKW = -1
  Private Const NM_FIRST = 0
  Private Const TTN_FIRST = -520
  Private Const TTN_LAST = -549
  Private Const WM_USER = &H400
  ' classname for buttons
  Private Const WC_TOOLTIP = "tooltips_class32"

  ' common controls messages
  Private Const CCM_FIRST = &H2000
  Private Const CCM_DPISCALE = (CCM_FIRST + &HC)
  Private Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)
  Private Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
  Private Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
  Private Const CCM_GETVERSION = (CCM_FIRST + &H8)
  Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
  Private Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)
  Private Const CCM_SETNOTIFYWINDOW = (CCM_FIRST + &H9)
  Private Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
  Private Const CCM_SETVERSION = (CCM_FIRST + &H7)
  Private Const CCM_SETWINDOWTHEME = (CCM_FIRST + &HB)

  ' constants used with CustomDraw
  Private Const CDDS_MAPPART = &H5
  Private Const CDDS_POSTERASE = &H4
  Private Const CDDS_PREERASE = &H3
  Private Const CDDS_ITEM = &H10000
  Private Const CDDS_PREPAINT = &H1
  Private Const CDDS_POSTPAINT = &H2
  Private Const CDDS_ITEMPOSTERASE = (CDDS_ITEM Or CDDS_POSTERASE)
  Private Const CDDS_ITEMPOSTPAINT = (CDDS_ITEM Or CDDS_POSTPAINT)
  Private Const CDDS_ITEMPREERASE = (CDDS_ITEM Or CDDS_PREERASE)
  Private Const CDDS_ITEMPREPAINT = (CDDS_ITEM Or CDDS_PREPAINT)
  Private Const CDDS_SUBITEM = &H20000
  Private Const CDIS_CHECKED = &H8
  Private Const CDIS_DEFAULT = &H20
  Private Const CDIS_DISABLED = &H4
  Private Const CDIS_FOCUS = &H10
  Private Const CDIS_GRAYED = &H2
  Private Const CDIS_HOT = &H40
  Private Const CDIS_INDETERMINATE = &H100
  Private Const CDIS_MARKED = &H80
  Private Const CDIS_SELECTED = &H1
  Private Const CDIS_SHOWKEYBOARDCUES = &H200
  Private Const CDRF_DODEFAULT = &H0
  Private Const CDRF_NEWFONT = &H2
  Private Const CDRF_NOTIFYITEMDRAW = &H20
  Private Const CDRF_NOTIFYPOSTERASE = &H40
  Private Const CDRF_NOTIFYPOSTPAINT = &H10
  Private Const CDRF_NOTIFYSUBITEMDRAW = &H20
  Private Const CDRF_SKIPDEFAULT = &H4

  ' tooltip styles
  Private Const TTS_ALWAYSTIP = &H1
  Private Const TTS_BALLOON = &H40
  Private Const TTS_CLOSE = &H80
  Private Const TTS_NOANIMATE = &H10
  Private Const TTS_NOFADE = &H20
  Private Const TTS_NOPREFIX = &H2
  Private Const TTS_USEVISUALSTYLE = &H100     ' requires Vista

  ' tooltip messages
  Private Const TTM_ACTIVATE = (WM_USER + 1)
  Private Const TTM_ADDTOOLA = (WM_USER + 4)
  Private Const TTM_ADDTOOLW = (WM_USER + 50)
  Private Const TTM_ADJUSTRECT = (WM_USER + 31)
  Private Const TTM_DELTOOLA = (WM_USER + 5)
  Private Const TTM_DELTOOLW = (WM_USER + 51)
  Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
  Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
  Private Const TTM_GETBUBBLESIZE = (WM_USER + 30)
  Private Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
  Private Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
  Private Const TTM_GETDELAYTIME = (WM_USER + 21)
  Private Const TTM_GETMARGIN = (WM_USER + 27)
  Private Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
  Private Const TTM_GETTEXTA = (WM_USER + 11)
  Private Const TTM_GETTEXTW = (WM_USER + 56)
  Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
  Private Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
  Private Const TTM_GETTITLE = (WM_USER + 35)
  Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
  Private Const TTM_GETTOOLINFOA = (WM_USER + 8)
  Private Const TTM_GETTOOLINFOW = (WM_USER + 53)
  Private Const TTM_HITTESTA = (WM_USER + 10)
  Private Const TTM_HITTESTW = (WM_USER + 55)
  Private Const TTM_NEWTOOLRECTA = (WM_USER + 6)
  Private Const TTM_NEWTOOLRECTW = (WM_USER + 52)
  Private Const TTM_POP = (WM_USER + 28)
  Private Const TTM_POPUP = (WM_USER + 34)
  Private Const TTM_RELAYEVENT = (WM_USER + 7)
  Private Const TTM_SETDELAYTIME = (WM_USER + 3)
  Private Const TTM_SETMARGIN = (WM_USER + 26)
  Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
  Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
  Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
  Private Const TTM_SETTITLEA = (WM_USER + 32)
  Private Const TTM_SETTITLEW = (WM_USER + 33)
  Private Const TTM_SETTOOLINFOA = (WM_USER + 9)
  Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
  Private Const TTM_SETWINDOWTHEME = CCM_SETWINDOWTHEME
  Private Const TTM_TRACKACTIVATE = (WM_USER + 17)
  Private Const TTM_TRACKPOSITION = (WM_USER + 18)
  Private Const TTM_UPDATE = (WM_USER + 29)
  Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
  Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
  Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16)

  ' common controls notifications
  Private Const NM_CHAR = (NM_FIRST - 18)
  Private Const NM_CLICK = (NM_FIRST - 2)
  Private Const NM_CUSTOMDRAW = (NM_FIRST - 12)
  Private Const NM_DBLCLK = (NM_FIRST - 3)
  Private Const NM_HOVER = (NM_FIRST - 13)
  Private Const NM_KEYDOWN = (NM_FIRST - 15)
  Private Const NM_KILLFOCUS = (NM_FIRST - 8)
  Private Const NM_LDOWN = (NM_FIRST - 20)
  Private Const NM_NCHITTEST = (NM_FIRST - 14)
  Private Const NM_OUTOFMEMORY = (NM_FIRST - 1)
  Private Const NM_RCLICK = (NM_FIRST - 5)
  Private Const NM_RDBLCLK = (NM_FIRST - 6)
  Private Const NM_RDOWN = (NM_FIRST - 21)
  Private Const NM_RELEASEDCAPTURE = (NM_FIRST - 16)
  Private Const NM_RETURN = (NM_FIRST - 4)
  Private Const NM_SETCURSOR = (NM_FIRST - 17)
  Private Const NM_SETFOCUS = (NM_FIRST - 7)
  Private Const NM_THEMECHANGED = (NM_FIRST - 22)
  Private Const NM_TOOLTIPSCREATED = (NM_FIRST - 19)

  ' tooltip notifications
  Private Const TTN_GETDISPINFOA = (TTN_FIRST - 0)
  Private Const TTN_GETDISPINFOW = (TTN_FIRST - 10)
  Private Const TTN_LINKCLICK = (TTN_FIRST - 3)
  Private Const TTN_NEEDTEXTA = TTN_GETDISPINFOA
  Private Const TTN_NEEDTEXTW = TTN_GETDISPINFOW
  Private Const TTN_POP = (TTN_FIRST - 2)
  Private Const TTN_SHOW = (TTN_FIRST - 1)

  ' tooltip flags
  Private Const TTF_ABSOLUTE = &H80
  Private Const TTF_RTLREADING = &H4
  Private Const TTF_CENTERTIP = &H2
  Private Const TTF_DI_SETITEM = &H8000
  Private Const TTF_IDISHWND = &H1
  Private Const TTF_PARSELINKS = &H1000
  Private Const TTF_SUBCLASS = &H10
  Private Const TTF_TRACK = &H20
  Private Const TTF_TRANSPARENT = &H100

  ' constants for TTM_GETDELAYTIME
  Private Const TTDT_AUTOMATIC = 0
  Private Const TTDT_AUTOPOP = 2
  Private Const TTDT_INITIAL = 3
  Private Const TTDT_RESHOW = 1

  ' constants for TTM_SETTITLE
  Private Const TTI_ERROR = 3
  Private Const TTI_INFO = 1
  Private Const TTI_NONE = 0
  Private Const TTI_WARNING = 2


  Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
  End Type

  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
  End Type

  Private Type NMTTCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    uDrawFlags As Long
  End Type

  Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
  End Type

  Private Type Size
    cx As Long
    cy As Long
  End Type

  Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    RECT As RECT
    hinst As Long
    lpszText As Long
    lParam As Long
    lpReserved As Long     ' requires Windows XP or higher
  End Type

  Private Type POINT
    x As Long
    y As Long
  End Type

  Private Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    lTime As Long
    pt As Long
  End Type

  Private Type NMCHAR
    hdr As NMHDR
    ch As Long
    dwItemPrev As Long
    dwItemNext As Long
  End Type

  Private Type NMKEY
    hdr As NMHDR
    nVKey As Long
    uFlags As Long
  End Type

  Private Type NMMOUSE
    hdr As NMHDR
    dwItemSpec As Long
    dwItemData As Long
    pt As POINT
    dwHitInfo As POINT
  End Type

  Private Type NMOBJECTNOTIFY
    hdr As NMHDR
    IItem As Long
    pObject As Long
    hResult As Long
    dwFlags As Long
  End Type

  Private Type NMTOOLTIPSCREATED
    hdr As NMHDR
    hwndToolTips As Long
  End Type

  Private Type TT_HITTESTINFOA
    hWnd As Long
    pt As POINT
    ti As TOOLINFO
  End Type

  Private Type TT_HITTESTINFOW
    hWnd As Long
    pt As POINT
    ti As TOOLINFO
  End Type

  Private Type INITCOMMONCONTROLSEXData
    dwSize As Long
    dwICC As Long
  End Type

  Private Type NMTTDISPINFOA
    hdr As NMHDR
    lpszText As Long
    szText(80) As Byte
    hinst As Long
    uFlags As Long
    lParam As Long
  End Type

  Private Type NMTTDISPINFOW
    hdr As NMHDR
    lpszText As Long
    szText(160) As Byte
    hinst As Long
    uFlags As Long
    lParam As Long
  End Type

  Private Type TTGETTITLE
    dwSize As Long
    uTitleBitmap As Long
    cch As Long
    pszTitle As Long
  End Type


  Private bAttached As Boolean
  Private bComctl32Version600OrNewer As Boolean
  #If CustomDraw Then
    Private BMPSize As Size
  #End If
  #If NeedsSubclassing Then
    Private pASMWrapper As Long
  #End If
  Private oldWinMain As Long
  Private oldWinMainParent As Long
  #If CustomDraw Then
    Private rcContent As RECT
    Private TextSize As Size
  #End If
  Private TOOLINFOSize As Long
  #If FullFeatured Then
    Private tooltipText As String
  #End If

  #If FullFeatured Then
    Private propAlwaysTip As Boolean
    Private propAutoPopDelayTime As Integer
    Private propBackColor As OLE_COLOR
    Private propBalloonStyle As Boolean
  #End If
  #If CustomDraw Then
    Private propBitmapMarginX As Long
    Private propBitmapMarginY As Long
    Private propBitmapPosition As BitmapPositionConstants
    Private prophBitmap As Long
  #End If
  #If FullFeatured Then
    Private prophFont As Long
  #End If
  Private prophWnd As Long
  Private prophWndParent As Long
  #If FullFeatured Then
    Private propInitialDelayTime As Long
    Private propMargin As RECT
  #End If
  #If CustomDraw Then
    Private propMaskColor As OLE_COLOR
  #End If
  #If FullFeatured Then
    Private propMaxTipWidth As Long
    Private propNoAnimation As Boolean
    Private propNoFade As Boolean
    Private propReshowDelayTime As Long
    Private propShowCloseButton As Boolean
    Private propUseVisualStyle As Boolean
  #End If
  Private propText As String
  #If FullFeatured Then
    Private propTextColor As OLE_COLOR
    Private propTitleIcon As ToolTipTitleIconConstants
    Private propTitleText As String
  #End If
  #If CustomDraw Then
    Private propUseMaskColor As Boolean
  #End If


  #If CustomDraw Then
    Private Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
    Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
  #End If
  #If NeedsSubclassing Then
    #If Unicode Then
      Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
    #Else
      Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    #End If
    Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
  #End If
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal cb As Long)
  #If CustomDraw Then
    Private Declare Function CreateBitmapAsLong Lib "gdi32.dll" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByVal lpvBits As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  #End If
  #If Unicode Then
    Private Declare Function CreateWindowExAsLong Lib "user32.dll" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long
  #Else
    Private Declare Function CreateWindowExAsLong Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long
  #End If
  #If CustomDraw Then
    Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  #End If
  Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (pdvi As DLLVERSIONINFO) As Long
  #If CustomDraw Then
    Private Declare Function EndPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
    Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lprc As RECT, ByVal hbr As Long) As Long
    Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
    #If Unicode Then
      Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectW" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, lpvObject As Any) As Long
    #Else
      Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, lpvObject As Any) As Long
    #End If
  #End If
  Private Declare Function GetMessagePos Lib "user32.dll" () As Long
  Private Declare Function GetMessageTime Lib "user32.dll" () As Long
  Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
  #If FullFeatured Then
    #If Unicode Then
      Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    #Else
      Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    #End If
  #End If
  Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEXData) As Long
  #If FullFeatured Then
    Private Declare Function lstrcpynAAsLong Lib "kernel32.dll" Alias "lstrcpynA" (ByVal lpString1 As Long, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
    Private Declare Function lstrcpynWAsLong Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
  #End If
  Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
  #If FullFeatured Then
    Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal clr As Long, ByVal hPal As Long, pcolorref As Long) As Long
    #If Unicode Then
      Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    #Else
      Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    #End If
  #End If
  #If CustomDraw Then
    Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hgdiobj As Long) As Long
  #End If
  #If Unicode Then
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  #Else
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  #End If
  #If CustomDraw Then
    Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  #End If
  Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
  #If Unicode Then
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  #Else
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  #End If
  #If CustomDraw Then
    Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
  #End If
  Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


  #If FullFeatured Then
    Public Event BeforeHide()
    Public Event BeforeShow()
    Public Event ClickedLink()
    Public Event NeedText(ByRef TextToDisplay As String, ByRef StoreData As Boolean)
    Public Event ProcessMessage(ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal callOldWinMain As Boolean, ByVal ValueToReturn As Long)
  #End If


Private Sub Class_Initialize()
  Const ICC_BAR_CLASSES = &H4
  #If NeedsSubclassing Then
    Const MEM_COMMIT = &H1000
    Const PAGE_EXECUTE_READWRITE = &H40
    Dim asm(0 To 103) As Byte
    Dim pRecCounter As Long     ' recursion counter to make recursions work
    Dim pDeleteFlag As Long     ' if this flag is set, the wrapper will delete itself
    Dim pThis As Long
    Dim pVirtualFree As Long
    Dim pVTable As Long
    Dim pWinMain As Long
  #End If
  Dim Data As INITCOMMONCONTROLSEXData
  Dim DLLVerData As DLLVERSIONINFO
  Dim tmp As TOOLINFO

  Data.dwICC = ICC_BAR_CLASSES
  Data.dwSize = LenB(Data)
  InitCommonControlsEx Data

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  If bComctl32Version600OrNewer Then
    TOOLINFOSize = LenB(tmp)
  Else
    TOOLINFOSize = LenB(tmp) - LenB(tmp.lpReserved)
  End If

  #If NeedsSubclassing Then
    ' this is a trick to simulate AddressOf for methods that belong to an object
    pASMWrapper = VirtualAlloc(0, 104, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pASMWrapper Then
      pVirtualFree = GetProcAddress(GetModuleHandle(StrPtr("kernel32.dll")), "VirtualFree")
      pRecCounter = pASMWrapper + 96     ' address of the recursion counter variable
      pDeleteFlag = pASMWrapper + 100     ' address of the delete flag

      ' get the VTable-Ptr
      pThis = ObjPtr(Me)
      CopyMemory VarPtr(pVTable), pThis, LenB(pVTable)
      ' get the pointer to the first public procedure (must be our WinMain)
      CopyMemory VarPtr(pWinMain), pVTable + &H1C, LenB(pWinMain)

      asm(0) = &H90     ' NOP
      'asm(0) = &HCC     ' INT 3 (software interrupt for debuging)
      asm(1) = &HFF     ' INC the recursion counter
      asm(2) = &H5
      CopyMemory VarPtr(asm(3)), VarPtr(pRecCounter), LenB(pRecCounter)
      asm(7) = &H6A     ' PUSH 0
      asm(8) = &H0
      asm(9) = &H54     ' PUSH ESP
      
      ' lay the addresses of the WinMain's arguments onto the stack
      asm(10) = &HFF    ' PUSH [ESP+18h] (lParam)
      asm(11) = &H74
      asm(12) = &H24
      asm(13) = &H18
      asm(14) = &HFF    ' PUSH [ESP+18h] (wParam)
      asm(15) = &H74
      asm(16) = &H24
      asm(17) = &H18
      asm(18) = &HFF    ' PUSH [ESP+18h] (msg)
      asm(19) = &H74
      asm(20) = &H24
      asm(21) = &H18
      asm(22) = &HFF    ' PUSH [ESP+18h] (hWnd)
      asm(23) = &H74
      asm(24) = &H24
      asm(25) = &H18
      ' lay the current object's pointer onto the stack
      asm(26) = &H68     ' PUSH 32-bit-immediate
      CopyMemory VarPtr(asm(27)), VarPtr(pThis), LenB(pVTable)

      asm(31) = &HB8     ' MOV EAX, (address of our WinProc)
      CopyMemory VarPtr(asm(32)), VarPtr(pWinMain), LenB(pWinMain)
      ' call our WinMain
      asm(36) = &HFF     ' CALL EAX
      asm(37) = &HD0
      asm(38) = &HFF     ' DEC the recursion counter
      asm(39) = &HD
      CopyMemory VarPtr(asm(40)), VarPtr(pRecCounter), LenB(pRecCounter)
      asm(44) = &HA1     ' MOV EAX, (delete flag)
      CopyMemory VarPtr(asm(45)), VarPtr(pDeleteFlag), LenB(pDeleteFlag)
      asm(49) = &H85     ' TEST EAX, EAX
      asm(50) = &HC0
      asm(51) = &H75     ' JNE
      asm(52) = &H4
      asm(53) = &H58     ' POP EAX (return value)
      asm(54) = &HC2     ' RET 10h
      asm(55) = &H10
      asm(56) = &H0
      asm(57) = &HA1     ' MOV EAX (our recursion counter)
      CopyMemory VarPtr(asm(58)), VarPtr(pRecCounter), LenB(pRecCounter)
      asm(62) = &H85     ' TEST EAX, EAX
      asm(63) = &HC0
      asm(64) = &H74     ' JE
      asm(65) = &H4
      asm(66) = &H58     ' POP EAX (return value)
      asm(67) = &HC2     ' RET 10h
      asm(68) = &H10
      asm(69) = &H0
      asm(70) = &H58     ' POP EAX (return value)
      asm(71) = &H59     ' POP ECX (return address)
      asm(72) = &H58     ' POP EAX (hWnd)
      asm(73) = &H58     ' POP EAX (msg)
      asm(74) = &H58     ' POP EAX (wParam)
      asm(75) = &H58     ' POP EAX (lParam)
      asm(76) = &H68     ' PUSH MEM_RELEASE (8000h)
      asm(77) = &H0
      asm(78) = &H80
      asm(79) = &H0
      asm(80) = &H0
      asm(81) = &H6A     ' PUSH 0
      asm(82) = &H0
      asm(83) = &H68     ' PUSH address of the wrapper
      CopyMemory VarPtr(asm(84)), VarPtr(pASMWrapper), LenB(pASMWrapper)
      asm(88) = &H51     ' PUSH ECX (return address)
      asm(89) = &HB8     ' MOV EAX (address of VirtualFree)
      CopyMemory VarPtr(asm(90)), VarPtr(pVirtualFree), LenB(pVirtualFree)
      asm(94) = &HFF     ' JMP EAX
      asm(95) = &HE0
      asm(96) = &H0      ' our recursion counter sits here
      asm(97) = &H0
      asm(98) = &H0
      asm(99) = &H0
      asm(100) = &H0     ' our delete flag sits here
      asm(101) = &H0
      asm(102) = &H0
      asm(103) = &H0

      CopyMemory pASMWrapper, VarPtr(asm(0)), 104
    End If
  #End If

  #If FullFeatured Then
    propAlwaysTip = True
    propAutoPopDelayTime = 5000
    propBackColor = SystemColorConstants.vbInfoBackground
    propBalloonStyle = False
  #End If
  #If CustomDraw Then
    propBitmapMarginX = 7
    propBitmapMarginY = 7
    propBitmapPosition = BitmapPositionConstants.bpLeft
  #End If
  #If FullFeatured Then
    propInitialDelayTime = 500
  #End If
  #If CustomDraw Then
    propMaskColor = ColorConstants.vbMagenta
  #End If
  #If FullFeatured Then
    propMaxTipWidth = 300
    propNoAnimation = False
    propNoFade = False
    propReshowDelayTime = 100
    propShowCloseButton = False
    propUseVisualStyle = True
    propTextColor = SystemColorConstants.vbInfoText
    propTitleIcon = ToolTipTitleIconConstants.tttiNone
    propTitleText = ""
  #End If
  #If CustomDraw Then
    propUseMaskColor = True
  #End If
End Sub

Private Sub Class_Terminate()
  #If NeedsSubclassing Then
    Const MEM_RELEASE = &H8000&
    Dim deleteFlag As Long
    Dim recCounter As Long     ' recursion counter
  #End If

  If bAttached Then
    Me.Detach
  Else
    Me.Destroy
  End If

  #If NeedsSubclassing Then
    If pASMWrapper Then
      ' check our ref counter
      CopyMemory VarPtr(recCounter), pASMWrapper + 96, LenB(recCounter)
      If recCounter = 0 Then
        VirtualFree pASMWrapper, 0, MEM_RELEASE
      Else
        ' the wrapper is within a recursion and must delete itself
        deleteFlag = 1
        CopyMemory pASMWrapper + 100, VarPtr(deleteFlag), LenB(deleteFlag)
      End If
    End If
  #End If
End Sub


#If NeedsSubclassing Then
  ' This must be the first public method!!
  ' this is *our* WinMain for the tooltip and its parent
  Public Function WinMain(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim ret As Long

    Select Case hWnd
      Case Me.hWnd
        ret = WinMain_ToolTip(hWnd, message, wParam, lParam)
      Case Me.hWndParent
        ret = WinMain_Parent(hWnd, message, wParam, lParam)
    End Select

    WinMain = ret
  End Function
#End If


#If FullFeatured Then
  Public Property Get AlwaysTip() As Boolean
    Const GWL_STYLE = (-16)
    Dim style As Long

    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      propAlwaysTip = ((style And TTS_ALWAYSTIP) = TTS_ALWAYSTIP)
    End If
    AlwaysTip = propAlwaysTip
  End Property

  Public Property Let AlwaysTip(ByVal newVal As Boolean)
    Const GWL_STYLE = (-16)
    Dim style As Long

    propAlwaysTip = newVal
    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      If propAlwaysTip Then
        style = style Or TTS_ALWAYSTIP
      Else
        style = style And Not TTS_ALWAYSTIP
      End If
      SetWindowLong Me.hWnd, GWL_STYLE, style
    End If
  End Property

  Public Property Get AutoPopDelayTime() As Integer
    If Me.hWnd Then
      propAutoPopDelayTime = SendMessageAsLong(Me.hWnd, TTM_GETDELAYTIME, TTDT_AUTOPOP, 0)
    End If
    AutoPopDelayTime = propAutoPopDelayTime
  End Property

  Public Property Let AutoPopDelayTime(ByVal newVal As Integer)
    propAutoPopDelayTime = newVal
    If Me.hWnd Then
      ' NOTE: For some reason the value won't be applied if we use SendMessage, so use PostMessage
      PostMessage Me.hWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, MAKELONG(propAutoPopDelayTime, 0)
    End If
  End Property

  Public Property Get BackColor() As OLE_COLOR
    If Me.hWnd Then
      propBackColor = SendMessage(Me.hWnd, TTM_GETTIPBKCOLOR, 0, 0)
    End If
    BackColor = propBackColor
  End Property

  Public Property Let BackColor(ByVal newVal As OLE_COLOR)
    propBackColor = newVal
    If Me.hWnd Then
      PostMessage Me.hWnd, TTM_SETTIPBKCOLOR, TranslateColor(propBackColor), 0
    End If
  End Property

  Public Property Get BalloonStyle() As Boolean
    Const GWL_STYLE = (-16)
    Dim style As Long

    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      propBalloonStyle = ((style And TTS_BALLOON) = TTS_BALLOON)
    End If
    BalloonStyle = propBalloonStyle
  End Property

  Public Property Let BalloonStyle(ByVal newVal As Boolean)
    Const GWL_STYLE = (-16)
    Const WS_BORDER = &H800000
    Dim style As Long

    propBalloonStyle = newVal
    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      If propBalloonStyle Then
        style = style Or TTS_BALLOON
        style = style And Not WS_BORDER
      Else
        style = style And Not TTS_BALLOON
        style = style Or WS_BORDER
      End If
      SetWindowLong Me.hWnd, GWL_STYLE, style
    End If
  End Property
#End If

#If CustomDraw Then
  Public Property Get BitmapMarginX() As Long
    BitmapMarginX = propBitmapMarginX
  End Property

  Public Property Let BitmapMarginX(ByVal newVal As Long)
    propBitmapMarginX = newVal
  End Property

  Public Property Get BitmapMarginY() As Long
    BitmapMarginY = propBitmapMarginY
  End Property

  Public Property Let BitmapMarginY(ByVal newVal As Long)
    propBitmapMarginY = newVal
  End Property

  Public Property Get BitmapPosition() As BitmapPositionConstants
    BitmapPosition = propBitmapPosition
  End Property

  Public Property Let BitmapPosition(ByVal newVal As BitmapPositionConstants)
    propBitmapPosition = newVal
  End Property

  Public Property Get hBitmap() As Long
    hBitmap = prophBitmap
  End Property

  Public Property Let hBitmap(ByVal newVal As Long)
    Dim BMPData As BITMAP

    prophBitmap = newVal
    GetObjectAPI prophBitmap, LenB(BMPData), BMPData
    BMPSize.cx = BMPData.bmWidth
    BMPSize.cy = BMPData.bmHeight
  End Property
#End If

#If FullFeatured Then
  Public Property Get hFont() As Long
    Const WM_GETFONT = &H31

    If Me.hWnd Then
      prophFont = SendMessageAsLong(Me.hWnd, WM_GETFONT, 0, 0)
    End If
    hFont = prophFont
  End Property

  Public Property Let hFont(ByVal newVal As Long)
    Const WM_SETFONT = &H30

    prophFont = newVal
    If Me.hWnd Then
      PostMessage Me.hWnd, WM_SETFONT, prophFont, 1
    End If
  End Property
#End If

Public Property Get hWnd() As Long
  hWnd = prophWnd
End Property

Public Property Get hWndParent() As Long
  If Me.hWnd Then
    prophWndParent = GetParent(Me.hWnd)
  End If
  hWndParent = prophWndParent
End Property

Public Property Let hWndParent(ByVal newVal As Long)
  UnSubclassParent
  prophWndParent = newVal
  SubclassParent
  If Me.hWnd Then
    SetParent Me.hWnd, prophWndParent
  End If
End Property

#If FullFeatured Then
  Public Property Get InitialDelayTime() As Integer
    If Me.hWnd Then
      propInitialDelayTime = SendMessageAsLong(Me.hWnd, TTM_GETDELAYTIME, TTDT_INITIAL, 0)
    End If
    InitialDelayTime = propInitialDelayTime
  End Property

  Public Property Let InitialDelayTime(ByVal newVal As Integer)
    propInitialDelayTime = newVal
    If Me.hWnd Then
      ' NOTE: For some reason the value won't be applied if we use SendMessage, so use PostMessage
      PostMessage Me.hWnd, TTM_SETDELAYTIME, TTDT_INITIAL, MAKELONG(propInitialDelayTime, 0)
    End If
  End Property

  Public Property Get MarginBottom() As Long
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_GETMARGIN, 0, propMargin
    End If
    MarginBottom = propMargin.Bottom
  End Property

  Public Property Let MarginBottom(ByVal newVal As Long)
    propMargin.Bottom = newVal
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_SETMARGIN, 0, propMargin
    End If
  End Property

  Public Property Get MarginLeft() As Long
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_GETMARGIN, 0, propMargin
    End If
    MarginLeft = propMargin.Left
  End Property

  Public Property Let MarginLeft(ByVal newVal As Long)
    propMargin.Left = newVal
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_SETMARGIN, 0, propMargin
    End If
  End Property

  Public Property Get MarginRight() As Long
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_GETMARGIN, 0, propMargin
    End If
    MarginRight = propMargin.Right
  End Property

  Public Property Let MarginRight(ByVal newVal As Long)
    propMargin.Right = newVal
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_SETMARGIN, 0, propMargin
    End If
  End Property

  Public Property Get MarginTop() As Long
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_GETMARGIN, 0, propMargin
    End If
    MarginTop = propMargin.Top
  End Property

  Public Property Let MarginTop(ByVal newVal As Long)
    propMargin.Top = newVal
    If Me.hWnd Then
      SendMessage Me.hWnd, TTM_SETMARGIN, 0, propMargin
    End If
  End Property
#End If

#If CustomDraw Then
  Public Property Get MaskColor() As OLE_COLOR
    MaskColor = propMaskColor
  End Property

  Public Property Let MaskColor(ByVal newVal As OLE_COLOR)
    propMaskColor = newVal
  End Property
#End If

#If FullFeatured Then
  Public Property Get MaxTipWidth() As Long
    If Me.hWnd Then
      propMaxTipWidth = SendMessageAsLong(Me.hWnd, TTM_GETMAXTIPWIDTH, 0, 0)
    End If
    MaxTipWidth = propMaxTipWidth
  End Property

  Public Property Let MaxTipWidth(ByVal newVal As Long)
    propMaxTipWidth = newVal
    If Me.hWnd Then
      SendMessageAsLong Me.hWnd, TTM_SETMAXTIPWIDTH, 0, propMaxTipWidth
    End If
  End Property

  Public Property Get NoAnimation() As Boolean
    Const GWL_STYLE = (-16)
    Dim style As Long

    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      propNoAnimation = ((style And TTS_NOANIMATE) = TTS_NOANIMATE)
    End If
    NoAnimation = propNoAnimation
  End Property

  Public Property Let NoAnimation(ByVal newVal As Boolean)
    Const GWL_STYLE = (-16)
    Dim style As Long

    propNoAnimation = newVal
    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      If propNoAnimation Then
        style = style Or TTS_NOANIMATE
      Else
        style = style And Not TTS_NOANIMATE
      End If
      SetWindowLong Me.hWnd, GWL_STYLE, style
    End If
  End Property

  Public Property Get NoFade() As Boolean
    Const GWL_STYLE = (-16)
    Dim style As Long

    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      propNoFade = ((style And TTS_NOFADE) = TTS_NOFADE)
    End If
    NoFade = propNoFade
  End Property

  Public Property Let NoFade(ByVal newVal As Boolean)
    Const GWL_STYLE = (-16)
    Dim style As Long

    propNoFade = newVal
    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      If propNoFade Then
        style = style Or TTS_NOFADE
      Else
        style = style And Not TTS_NOFADE
      End If
      SetWindowLong Me.hWnd, GWL_STYLE, style
    End If
  End Property

  Public Property Get ReshowDelayTime() As Integer
    If Me.hWnd Then
      propReshowDelayTime = SendMessageAsLong(Me.hWnd, TTM_GETDELAYTIME, TTDT_RESHOW, 0)
    End If
    ReshowDelayTime = propReshowDelayTime
  End Property

  Public Property Let ReshowDelayTime(ByVal newVal As Integer)
    propReshowDelayTime = newVal
    If Me.hWnd Then
      ' NOTE: For some reason the value won't be applied if we use SendMessage, so use PostMessage
      PostMessage Me.hWnd, TTM_SETDELAYTIME, TTDT_RESHOW, MAKELONG(propReshowDelayTime, 0)
    End If
  End Property

  Public Property Get ShowCloseButton() As Boolean
    Const GWL_STYLE = (-16)
    Dim style As Long

    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      propShowCloseButton = ((style And TTS_CLOSE) = TTS_CLOSE)
    End If
    ShowCloseButton = propShowCloseButton
  End Property

  Public Property Let ShowCloseButton(ByVal newVal As Boolean)
    Const GWL_STYLE = (-16)
    Dim style As Long

    propShowCloseButton = newVal
    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      If propShowCloseButton Then
        style = style Or TTS_CLOSE
      Else
        style = style And Not TTS_CLOSE
      End If
      SetWindowLong Me.hWnd, GWL_STYLE, style
    End If
  End Property

  Public Property Get UseVisualStyle() As Boolean
    Const GWL_STYLE = (-16)
    Dim style As Long

    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      propUseVisualStyle = ((style And TTS_USEVISUALSTYLE) = TTS_USEVISUALSTYLE)
    End If
    UseVisualStyle = propUseVisualStyle
  End Property

  Public Property Let UseVisualStyle(ByVal newVal As Boolean)
    Const GWL_STYLE = (-16)
    Dim style As Long

    propUseVisualStyle = newVal
    If Me.hWnd Then
      style = GetWindowLong(Me.hWnd, GWL_STYLE)
      If propUseVisualStyle Then
        style = style Or TTS_USEVISUALSTYLE
      Else
        style = style And Not TTS_USEVISUALSTYLE
      End If
      SetWindowLong Me.hWnd, GWL_STYLE, style
    End If
  End Property
#End If

Public Property Get Text(ByVal hWndTool As Long) As String
  Dim Data As TOOLINFO

  If Me.hWnd Then
    With Data
      .cbSize = TOOLINFOSize
      .hinst = App.hInstance
      .hWnd = Me.hWndParent
      propText = String$(2048, Chr$(0))
      .lpszText = StrPtr(propText)
      .uId = hWndTool
      .uFlags = TTF_IDISHWND
    End With
    #If Unicode Then
      SendMessage Me.hWnd, TTM_GETTEXTW, 0, Data
    #Else
      SendMessage Me.hWnd, TTM_GETTEXTA, 0, Data
      propText = StrConv(propText, VbStrConv.vbUnicode)
    #End If
    propText = Left$(propText, lstrlenW(StrPtr(propText)))
  End If
  Text = propText
End Property

Public Property Let Text(ByVal hWndTool As Long, ByVal newVal As String)
  Dim Data As TOOLINFO

  If Me.Text(hWndTool) <> newVal Then     ' this check prevents flickering on Windows 7
    propText = newVal
    If Me.hWnd Then
      With Data
        .cbSize = TOOLINFOSize
        .hinst = App.hInstance
        .hWnd = Me.hWndParent
        .uId = hWndTool
        .uFlags = TTF_IDISHWND
        #If Unicode Then
          .lpszText = StrPtr(propText)
          SendMessage Me.hWnd, TTM_UPDATETIPTEXTW, 0, Data
        #Else
          .lpszText = StrPtr(StrConv(propText, VbStrConv.vbFromUnicode))
          SendMessage Me.hWnd, TTM_UPDATETIPTEXTA, 0, Data
        #End If
      End With
    End If
  End If
End Property

#If FullFeatured Then
  Public Property Get TextColor() As OLE_COLOR
    If Me.hWnd Then
      propTextColor = SendMessageAsLong(Me.hWnd, TTM_GETTIPTEXTCOLOR, 0, 0)
    End If
    TextColor = propTextColor
  End Property

  Public Property Let TextColor(ByVal newVal As OLE_COLOR)
    propTextColor = newVal
    If Me.hWnd Then
      PostMessage Me.hWnd, TTM_SETTIPTEXTCOLOR, TranslateColor(propTextColor), 0
    End If
  End Property

  Public Property Get TitleIcon() As ToolTipTitleIconConstants
    Dim Data As TTGETTITLE
    Dim str As String

    If Me.hWnd Then
      With Data
        .dwSize = LenB(Data)
        str = String$(1024, Chr$(0))
        .pszTitle = StrPtr(str)
        .cch = Len(str)

        SendMessage Me.hWnd, TTM_GETTITLE, 0, Data
        propTitleIcon = .uTitleBitmap
      End With
    End If
    TitleIcon = propTitleIcon
  End Property

  Public Property Let TitleIcon(ByVal newVal As ToolTipTitleIconConstants)
    propTitleIcon = newVal
    If Me.hWnd Then
      #If Unicode Then
        SendMessageAsLong Me.hWnd, TTM_SETTITLEW, propTitleIcon, StrPtr(propTitleText)
      #Else
        SendMessageAsLong Me.hWnd, TTM_SETTITLEA, propTitleIcon, StrPtr(StrConv(propTitleText, VbStrConv.vbFromUnicode))
      #End If
    End If
  End Property

  Public Property Get TitleText() As String
    Dim Data As TTGETTITLE
    Dim str As String

    If Me.hWnd Then
      With Data
        .dwSize = LenB(Data)
        str = String$(1024, Chr$(0))
        .pszTitle = StrPtr(str)
        .cch = Len(str)

        SendMessage Me.hWnd, TTM_GETTITLE, 0, Data
        propTitleText = Left$(str, lstrlenW(StrPtr(str)))
      End With
    End If
    TitleText = propTitleText
  End Property

  Public Property Let TitleText(ByVal newVal As String)
    propTitleText = newVal
    If Me.hWnd Then
      #If Unicode Then
        SendMessageAsLong Me.hWnd, TTM_SETTITLEW, propTitleIcon, StrPtr(propTitleText)
      #Else
        SendMessageAsLong Me.hWnd, TTM_SETTITLEA, propTitleIcon, StrPtr(StrConv(propTitleText, VbStrConv.vbFromUnicode))
      #End If
    End If
  End Property
#End If

#If CustomDraw Then
  Public Property Get UseMaskColor() As Boolean
    UseMaskColor = propUseMaskColor
  End Property

  Public Property Let UseMaskColor(ByVal newVal As Boolean)
    propUseMaskColor = newVal
  End Property
#End If


Public Sub Activate()
  If Me.hWnd Then
    SendMessageAsLong Me.hWnd, TTM_ACTIVATE, 1, 0
  End If
End Sub

#If FullFeatured Then
  Public Sub AddTool(ByVal hWndTool As Long, ByVal TextToDisplay As String, Optional ByVal SubClassTool As Boolean = True, Optional ByVal Transparent As Boolean = True, Optional ByVal AdjustToTool As Boolean = True, Optional ByVal ParseLinks As Boolean = True, Optional ByVal Track As Boolean = False)
#Else
  Public Sub AddTool(ByVal hWndTool As Long, ByVal TextToDisplay As String, Optional ByVal SubClassTool As Boolean = True, Optional ByVal Transparent As Boolean = True, Optional ByVal AdjustToTool As Boolean = True)
#End If
  Dim Data As TOOLINFO

  If Me.hWnd Then
    With Data
      .cbSize = TOOLINFOSize
      .hinst = App.hInstance
      .hWnd = Me.hWndParent
      .uId = hWndTool
      .uFlags = TTF_IDISHWND
      If AdjustToTool Then .uFlags = .uFlags Or TTF_CENTERTIP
      #If FullFeatured Then
        If ParseLinks Then .uFlags = .uFlags Or TTF_PARSELINKS
      #End If
      If SubClassTool Then .uFlags = .uFlags Or TTF_SUBCLASS
      #If FullFeatured Then
        If Track Then .uFlags = .uFlags Or TTF_TRACK Or TTF_ABSOLUTE
      #End If
      If Transparent Then .uFlags = .uFlags Or TTF_TRANSPARENT
      #If Unicode Then
        #If FullFeatured Then
          If TextToDisplay = "" Then
            .lpszText = LPSTR_TEXTCALLBACKW
          Else
            .lpszText = StrPtr(TextToDisplay)
          End If
        #Else
          .lpszText = StrPtr(TextToDisplay)
        #End If
        SendMessage Me.hWnd, TTM_ADDTOOLW, 0, Data
      #Else
        #If FullFeatured Then
          If TextToDisplay = "" Then
            .lpszText = LPSTR_TEXTCALLBACKA
          Else
            .lpszText = StrPtr(StrConv(TextToDisplay, VbStrConv.vbFromUnicode))
          End If
        #Else
          .lpszText = StrPtr(StrConv(TextToDisplay, VbStrConv.vbFromUnicode))
        #End If
        SendMessage Me.hWnd, TTM_ADDTOOLA, 0, Data
      #End If
    End With
  End If
End Sub

' attaches this object to an existing tooltip control
Public Sub Attach(ByVal hWndToolTip As Long)
  Me.Destroy
  prophWnd = hWndToolTip
  If Me.hWnd Then
    Subclass
    SubclassParent
  End If
  bAttached = True
End Sub

' creates a tooltip
#If FullFeatured Then
  Public Sub Create(Optional ByVal hWndParent As Long = -1, Optional ByVal hFont As Long = -1)
#Else
  Public Sub Create(Optional ByVal hWndParent As Long = -1)
#End If
  Const CW_USEDEFAULT = &H80000000
  Const WS_EX_TOOLWINDOW = &H80
  Const WS_POPUP = &H80000000
  Const WS_VISIBLE = &H10000000
  #If FullFeatured Then
    Const WM_GETFONT = &H31
    Const WM_SETFONT = &H30
  #End If
  Dim TTStyle As Long

  ' you'll never know...
  Me.Destroy
  If Me.hWnd = 0 Then
    If hWndParent <> -1 Then prophWndParent = hWndParent
    #If FullFeatured Then
      If hFont <> -1 Then
        Me.hFont = hFont
      ElseIf Me.hFont = -1 Then
        Me.hFont = SendMessageAsLong(Me.hWndParent, WM_GETFONT, 0, 0)
      End If
    #End If

    TTStyle = TTS_NOPREFIX
    #If FullFeatured Then
      If propAlwaysTip Then TTStyle = TTStyle Or TTS_ALWAYSTIP
      If propBalloonStyle Then TTStyle = TTStyle Or TTS_BALLOON
      If propNoAnimation Then TTStyle = TTStyle Or TTS_NOANIMATE
      If propNoFade Then TTStyle = TTStyle Or TTS_NOFADE
      If propShowCloseButton Then TTStyle = TTStyle Or TTS_CLOSE
      If propUseVisualStyle Then TTStyle = TTStyle Or TTS_USEVISUALSTYLE
    #End If

    ' create the window
    #If Unicode Then
      prophWnd = CreateWindowExAsLong(WS_EX_TOOLWINDOW, StrPtr(WC_TOOLTIP), 0, WS_POPUP Or WS_VISIBLE Or TTStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, Me.hWndParent, 0, App.hInstance, 0)
    #Else
      prophWnd = CreateWindowExAsLong(WS_EX_TOOLWINDOW, WC_TOOLTIP, vbNullString, WS_POPUP Or WS_VISIBLE Or TTStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, Me.hWndParent, 0, App.hInstance, 0)
    #End If
    If Me.hWnd = 0 Then GoTo MyError
  End If
  Subclass
  SubclassParent

  #If FullFeatured Then
    ' set the properties...
    PostMessage Me.hWnd, TTM_SETMAXTIPWIDTH, 0, propMaxTipWidth
    PostMessage Me.hWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, MAKELONG(propAutoPopDelayTime, 0)
    PostMessage Me.hWnd, TTM_SETDELAYTIME, TTDT_INITIAL, MAKELONG(propInitialDelayTime, 0)
    PostMessage Me.hWnd, TTM_SETDELAYTIME, TTDT_RESHOW, MAKELONG(propReshowDelayTime, 0)
    #If Unicode Then
      SendMessageAsLong Me.hWnd, TTM_SETTITLEW, propTitleIcon, StrPtr(propTitleText)
    #Else
      SendMessageAsLong Me.hWnd, TTM_SETTITLEA, propTitleIcon, StrPtr(StrConv(propTitleText, VbStrConv.vbFromUnicode))
    #End If

    ' ...and colors...
    PostMessage Me.hWnd, TTM_SETTIPBKCOLOR, TranslateColor(propBackColor), 0
    PostMessage Me.hWnd, TTM_SETTIPTEXTCOLOR, TranslateColor(propTextColor), 0

    ' ...and fonts
    PostMessage Me.hWnd, WM_SETFONT, prophFont, 1
  #End If
  Exit Sub

MyError:
End Sub

Public Sub Deactivate()
  If Me.hWnd Then
    SendMessageAsLong Me.hWnd, TTM_ACTIVATE, 0, 0
  End If
End Sub

Public Sub DelTool(ByVal hWndTool As Long)
  Dim Data As TOOLINFO

  If Me.hWnd Then
    With Data
      .cbSize = TOOLINFOSize
      .hWnd = Me.hWndParent
      .uId = hWndTool
    End With
    #If Unicode Then
      SendMessage Me.hWnd, TTM_DELTOOLW, 0, Data
    #Else
      SendMessage Me.hWnd, TTM_DELTOOLA, 0, Data
    #End If
  End If
End Sub

' destroys the frame
Public Sub Destroy()
  Const SW_HIDE = 0

  UnSubclassParent
  If prophWnd Then
    UnSubclass

    If Not bAttached Then
      ShowWindow prophWnd, SW_HIDE
      SetParent prophWnd, 0

      DestroyWindow prophWnd
      prophWnd = 0
    End If
  End If
End Sub

' detaches this object from the tooltip control (must have been previously attached)
Public Sub Detach()
  If bAttached Then
    UnSubclass
    UnSubclassParent
    prophWnd = 0

    bAttached = False
  End If
End Sub

#If FullFeatured Then
  Public Sub Hide(Optional ByVal hWndTool As Long = 0)
    Dim Data As TOOLINFO

    If Me.hWnd Then
      If hWndTool Then
        With Data
          .cbSize = TOOLINFOSize
          .hWnd = Me.hWndParent
          .uId = hWndTool
          .uFlags = TTF_IDISHWND
        End With
        SendMessage Me.hWnd, TTM_TRACKACTIVATE, 0, Data
      Else
        SendMessageAsLong Me.hWnd, TTM_POP, 0, 0
      End If
    End If
  End Sub
#End If

Public Sub RelayMouseEvent(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long)
  Dim messageDetails As Msg

  If Me.hWnd Then
    Me.Activate
    messageDetails.hWnd = hWnd
    messageDetails.message = message
    messageDetails.wParam = wParam
    messageDetails.lParam = lParam
    SendMessage Me.hWnd, TTM_RELAYEVENT, 0, messageDetails
  End If
End Sub

#If FullFeatured Then
  Public Sub Show(ByVal hWndTool As Long)
    Dim Data As TOOLINFO

    If Me.hWnd Then
      With Data
        .cbSize = TOOLINFOSize
        .hWnd = Me.hWndParent
        .uId = hWndTool
        .uFlags = TTF_IDISHWND
      End With
      SendMessage Me.hWnd, TTM_TRACKACTIVATE, 1, Data
    End If
  End Sub

  Public Sub TrackTo(ByVal x As Long, ByVal y As Long)
    If Me.hWnd Then
      SendMessageAsLong Me.hWnd, TTM_TRACKPOSITION, 0, MAKELONG(x, y)
    End If
  End Sub
#End If


#If CustomDraw Then
  Private Sub DrawTransparentBitmap(ByVal hBitmap As Long, ByVal MaskClr As Long, ByVal hDC_Dest As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long)
    Const SRCAND = &H8800C6
    Const SRCCOPY = &HCC0020
    Const SRCINVERT = &H660046
    Dim hBMP As Long
    Dim hBMP_Mask As Long
    Dim hBMP_Mask_Old As Long
    Dim hBMP_Old As Long
    Dim hBMP_Pic_Old As Long
    Dim hBMP_PicCopy As Long
    Dim hBMP_PicCopy_Old As Long
    Dim hDC_Mem As Long
    Dim hDC_Mem_Mask As Long
    Dim hDC_Mem_Pic As Long
    Dim hDC_Mem_PicCopy As Long

    hDC_Mem = CreateCompatibleDC(hDC_Dest)     ' our working DC
    If hDC_Mem Then
      hBMP = CreateCompatibleBitmap(hDC_Dest, cx, cy)
      If hBMP Then
        hBMP_Old = SelectObject(hDC_Mem, hBMP)
        hDC_Mem_Mask = CreateCompatibleDC(hDC_Dest)     ' our working DC for the mask
        If hDC_Mem_Mask Then
          hBMP_Mask = CreateBitmapAsLong(cx, cy, 1, 1, 0)
          If hBMP_Mask Then
            hBMP_Mask_Old = SelectObject(hDC_Mem_Mask, hBMP_Mask)
            hDC_Mem_Pic = CreateCompatibleDC(hDC_Dest)     ' the DC that holds <hBitmap>
            If hDC_Mem_Pic Then
              hBMP_Pic_Old = SelectObject(hDC_Mem_Pic, hBitmap)
              hDC_Mem_PicCopy = CreateCompatibleDC(hDC_Dest)     ' a DC that holds a "fresh" copy of <hBitmap>
              If hDC_Mem_PicCopy Then
                hBMP_PicCopy = CreateCompatibleBitmap(hDC_Dest, cx, cy)
                If hBMP_PicCopy Then
                  hBMP_PicCopy_Old = SelectObject(hDC_Mem_PicCopy, hBMP_PicCopy)

                  ' copy the target's background into our working DC
                  BitBlt hDC_Mem, 0, 0, cx, cy, hDC_Dest, x, y, SRCCOPY

                  ' create a copy of the bitmap
                  BitBlt hDC_Mem_PicCopy, 0, 0, cx, cy, hDC_Mem_Pic, 0, 0, SRCCOPY

                  ' combine the bitmap several times to get a mask
                  SetBkColor hDC_Mem_PicCopy, MaskClr
                  BitBlt hDC_Mem_Mask, 0, 0, cx, cy, hDC_Mem_PicCopy, 0, 0, SRCCOPY
                  BitBlt hDC_Mem_PicCopy, 0, 0, cx, cy, hDC_Mem_Mask, 0, 0, SRCINVERT
                  BitBlt hDC_Mem, 0, 0, cx, cy, hDC_Mem_Mask, 0, 0, SRCAND
                  BitBlt hDC_Mem, 0, 0, cx, cy, hDC_Mem_PicCopy, 0, 0, SRCINVERT

                  ' transfer the result to the source DC
                  BitBlt hDC_Dest, x, y, cx, cy, hDC_Mem, 0, 0, SRCCOPY

                  SelectObject hDC_Mem_PicCopy, hBMP_PicCopy_Old
                  DeleteObject hBMP_PicCopy
                End If
                DeleteDC hDC_Mem_PicCopy
              End If
              SelectObject hDC_Mem_Pic, hBMP_Pic_Old
              DeleteDC hDC_Mem_Pic
            End If
            SelectObject hDC_Mem_Mask, hBMP_Mask_Old
            DeleteObject hBMP_Mask
          End If
          DeleteDC hDC_Mem_Mask
        End If
        SelectObject hDC_Mem, hBMP_Old
        DeleteObject hBMP
      End If
      DeleteDC hDC_Mem
    End If
  End Sub

  Private Function HandleWM_PAINT(ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const PRF_CHILDREN = &H10&
    Const PRF_CLIENT = &H4&
    Const PRF_OWNED = &H20&
    Const SRCCOPY = &HCC0020
    Const WM_PRINTCLIENT = &H318
    Dim Data As PAINTSTRUCT
    Dim hBMP_Content As Long
    Dim hBMP_Content_Old As Long
    Dim hBMP_Old As Long
    Dim hBMP_Text As Long
    Dim hBMP_Text_Old As Long
    Dim hBrBk As Long
    Dim hDC_Mem_BMP As Long
    Dim hDC_Mem_Content As Long
    Dim hDC_Mem_Text As Long
    Dim hDC_TT As Long
    Dim rcRedraw As RECT
    Dim rcText As RECT
    Dim x As Long
    Dim y As Long
    Static bRecursive As Boolean

    If bRecursive Then Exit Function

    bRecursive = True
    BeginPaint hWnd, Data

    hDC_TT = Data.hDC
    rcRedraw = Data.rcPaint

    hDC_Mem_Content = CreateCompatibleDC(hDC_TT)
    If hDC_Mem_Content Then
      With rcContent
        hBMP_Content = CreateCompatibleBitmap(hDC_TT, .Right - .Left, .Bottom - .Top)
      End With
      If hBMP_Content Then
        hBMP_Content_Old = SelectObject(hDC_Mem_Content, hBMP_Content)

        #If FullFeatured Then
          hBrBk = CreateSolidBrush(Me.BackColor)
        #Else
          hBrBk = CreateSolidBrush(SendMessageAsLong(Me.hWnd, TTM_GETTIPBKCOLOR, 0, 0))
        #End If
        If hBrBk Then FillRect hDC_Mem_Content, rcContent, hBrBk

        If Me.UseMaskColor Then
          hDC_Mem_BMP = -1
        Else
          hDC_Mem_BMP = CreateCompatibleDC(hDC_TT)
        End If
        If hDC_Mem_BMP Then
          If hDC_Mem_BMP <> -1 Then hBMP_Old = SelectObject(hDC_Mem_BMP, Me.hBitmap)

          hDC_Mem_Text = CreateCompatibleDC(hDC_TT)
          If hDC_Mem_Text Then
            hBMP_Text = CreateCompatibleBitmap(hDC_TT, TextSize.cx, TextSize.cy)
            If hBMP_Text Then
              hBMP_Text_Old = SelectObject(hDC_Mem_Text, hBMP_Text)

              rcText.Left = 0
              rcText.Right = rcText.Left + TextSize.cx
              rcText.Top = 0
              rcText.Bottom = rcText.Top + TextSize.cy
              If hBrBk Then FillRect hDC_Mem_Text, rcText, hBrBk

              ' let Windows draw the actual tooltip into our memory DC
              If oldWinMain Then
                CallWindowProc oldWinMain, hWnd, WM_PRINTCLIENT, hDC_Mem_Text, PRF_CHILDREN Or PRF_CLIENT Or PRF_OWNED
              End If

              ' now put it all together into the memory DC
              ' start with the bitmap
              Select Case Me.BitmapPosition
                Case BitmapPositionConstants.bpLeft
                  x = Me.BitmapMarginX
                  y = ((rcContent.Bottom - rcContent.Top) - BMPSize.cy) / 2
                Case BitmapPositionConstants.bpTop
                  x = ((rcContent.Right - rcContent.Left) - BMPSize.cx) / 2
                  y = Me.BitmapMarginY
                Case BitmapPositionConstants.bpRight
                  x = TextSize.cx
                  #If FullFeatured Then
                    If Me.TitleText = "" Then x = x + Me.BitmapMarginX
                  #Else
                    x = x + Me.BitmapMarginX
                  #End If
                  y = ((rcContent.Bottom - rcContent.Top) - BMPSize.cy) / 2
                Case BitmapPositionConstants.bpBottom
                  x = ((rcContent.Right - rcContent.Left) - BMPSize.cx) / 2
                  y = TextSize.cy
                  #If FullFeatured Then
                    If Me.TitleText = "" Then y = y + Me.BitmapMarginY
                  #Else
                    y = y + Me.BitmapMarginY
                  #End If
              End Select
              If Me.UseMaskColor Then
                DrawTransparentBitmap Me.hBitmap, Me.MaskColor, hDC_Mem_Content, x, y, BMPSize.cx, BMPSize.cy
              Else
                BitBlt hDC_Mem_Content, x, y, BMPSize.cx, BMPSize.cy, hDC_Mem_BMP, 0, 0, SRCCOPY
              End If

              ' continue with the actual tooltip
              Select Case Me.BitmapPosition
                Case BitmapPositionConstants.bpLeft
                  x = Me.BitmapMarginX + BMPSize.cx
                  #If FullFeatured Then
                    If Me.TitleText = "" Then x = x + Me.BitmapMarginX
                  #Else
                    x = x + Me.BitmapMarginX
                  #End If
                  y = ((rcContent.Bottom - rcContent.Top) - TextSize.cy) / 2
                Case BitmapPositionConstants.bpTop
                  x = ((rcContent.Right - rcContent.Left) - TextSize.cx) / 2
                  y = Me.BitmapMarginY + BMPSize.cy
                  #If FullFeatured Then
                    If Me.TitleText = "" Then y = y + Me.BitmapMarginY
                  #Else
                    y = y + Me.BitmapMarginY
                  #End If
                Case BitmapPositionConstants.bpRight
                  x = 0
                  #If FullFeatured Then
                    If Me.TitleText = "" Then x = 3
                  #Else
                    x = 3
                  #End If
                  y = ((rcContent.Bottom - rcContent.Top) - TextSize.cy) / 2
                Case BitmapPositionConstants.bpBottom
                  x = ((rcContent.Right - rcContent.Left) - TextSize.cx) / 2
                  y = 0
                  #If FullFeatured Then
                    If Me.TitleText = "" Then y = 3
                  #Else
                    y = 3
                  #End If
              End Select
              BitBlt hDC_Mem_Content, x, y, TextSize.cx, TextSize.cy, hDC_Mem_Text, 0, 0, SRCCOPY

              ' now copy the desired rectangle into the target DC
              With rcRedraw
                BitBlt hDC_TT, .Left, .Top, .Right - .Left, .Bottom - .Top, hDC_Mem_Content, .Left, .Top, SRCCOPY
              End With
              HandleWM_PAINT = 0

              SelectObject hDC_Mem_Text, hBMP_Text_Old
              DeleteObject hBMP_Text
            End If
            DeleteDC hDC_Mem_Text
          End If

          If hDC_Mem_BMP <> -1 Then
            SelectObject hDC_Mem_BMP, hBMP_Old
            DeleteDC hDC_Mem_BMP
          End If
        End If

        If hBrBk Then DeleteObject hBrBk
        SelectObject hDC_Mem_Content, hBMP_Content_Old
        DeleteObject hBMP_Content_Old
      End If
      DeleteDC hDC_Mem_Content
    End If

    EndPaint hWnd, Data
    bRecursive = False
  End Function
#End If

#If FullFeatured Then
  ' concatenates <Lo> and <Hi> to a 32-bit-number
  Private Function MAKELONG(ByVal Lo As Integer, ByVal Hi As Integer) As Long
    Dim ret As Long

    CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
    CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(Hi), LenB(Hi)

    MAKELONG = ret
  End Function
#End If

#If CustomDraw Then
  Private Sub ResizeToolTipWithBitmap()
    Const SWP_NOACTIVATE = &H10
    Const SWP_NOMOVE = &H2
    Const SWP_NOOWNERZORDER = &H200
    Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
    Const SWP_NOZORDER = &H4
    Dim rcClient As RECT
    Dim szNeeded As Size

    ' get size without bitmap
    GetClientRect Me.hWnd, rcClient
    TextSize.cx = rcClient.Right - rcClient.Left
    TextSize.cy = rcClient.Bottom - rcClient.Top

    ' calculate the desired size
    Select Case Me.BitmapPosition
      Case BitmapPositionConstants.bpLeft, BitmapPositionConstants.bpRight
        szNeeded.cx = Me.BitmapMarginX + BMPSize.cx + TextSize.cx
        #If FullFeatured Then
          If Me.TitleText = "" Then szNeeded.cx = szNeeded.cx + Me.BitmapMarginX
        #Else
          szNeeded.cx = szNeeded.cx + Me.BitmapMarginX
        #End If
        If (BMPSize.cy + 2 * Me.BitmapMarginY) > TextSize.cy Then
          szNeeded.cy = BMPSize.cy + 2 * Me.BitmapMarginY
        Else
          szNeeded.cy = TextSize.cy
        End If

      Case BitmapPositionConstants.bpTop, BitmapPositionConstants.bpBottom
        szNeeded.cy = Me.BitmapMarginY + BMPSize.cy + TextSize.cy
        #If FullFeatured Then
          If Me.TitleText = "" Then szNeeded.cy = szNeeded.cy + Me.BitmapMarginY
        #Else
          szNeeded.cy = szNeeded.cy + Me.BitmapMarginY
        #End If
        If (BMPSize.cx + 2 * Me.BitmapMarginX) > TextSize.cx Then
          szNeeded.cx = BMPSize.cx + 2 * Me.BitmapMarginX
        Else
          szNeeded.cx = TextSize.cx
        End If
    End Select

    ' don't forget we've a border of 1 pixel
    ' TODO: get border size from system settings
    szNeeded.cx = szNeeded.cx + 2
    szNeeded.cy = szNeeded.cy + 2

    ' apply the size
    SetWindowPos Me.hWnd, 0, 0, 0, szNeeded.cx, szNeeded.cy, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOREPOSITION Or SWP_NOZORDER
    GetClientRect Me.hWnd, rcContent
  End Sub
#End If

Private Sub Subclass()
  Const GWL_WNDPROC = (-4)

  #If NeedsSubclassing Then
    If oldWinMain = 0 Then
      If pASMWrapper Then
        oldWinMain = SetWindowLong(Me.hWnd, GWL_WNDPROC, pASMWrapper)
      End If
    End If
  #End If
End Sub

Private Sub SubclassParent()
  Const GWL_WNDPROC = (-4)

  #If NeedsSubclassing Then
    If oldWinMainParent = 0 Then
      If pASMWrapper Then
        oldWinMainParent = SetWindowLong(Me.hWndParent, GWL_WNDPROC, pASMWrapper)
      End If
    End If
  #End If
End Sub

#If FullFeatured Then
  ' transforms an OLE color to a RGB color
  Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional ByVal hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
  End Function
#End If

Private Sub UnSubclass()
  Const GWL_WNDPROC = (-4)

  If oldWinMain Then
    SetWindowLong Me.hWnd, GWL_WNDPROC, oldWinMain
    oldWinMain = 0
  End If
End Sub

Private Sub UnSubclassParent()
  Const GWL_WNDPROC = (-4)

  If oldWinMainParent Then
    SetWindowLong Me.hWndParent, GWL_WNDPROC, oldWinMainParent
    oldWinMainParent = 0
  End If
End Sub

#If NeedsSubclassing Then
  Private Function WinMain_Parent(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const NF_QUERY = 3
    Const NFR_ANSI = 1
    Const NFR_UNICODE = 2
    Const WM_NOTIFY = &H4E
    Const WM_NOTIFYFORMAT = &H55
    Dim callOldWinMain As Boolean
    #If FullFeatured Then
      Dim bStore As Boolean
      Dim DispDataA As NMTTDISPINFOA
      Dim DispDataW As NMTTDISPINFOW
    #End If
    Dim NotData As NMHDR
    Dim ret As Long

    callOldWinMain = True
    Select Case message
      Case WM_NOTIFY
        CopyMemory VarPtr(NotData), lParam, LenB(NotData)
        If NotData.hwndFrom = Me.hWnd Then
          Select Case NotData.code
            #If FullFeatured Then
              Case TTN_LINKCLICK
                RaiseEvent ClickedLink
              Case TTN_NEEDTEXTA
                tooltipText = ""
                RaiseEvent NeedText(tooltipText, bStore)
                If tooltipText <> "" Then
                  CopyMemory VarPtr(DispDataA), lParam, LenB(DispDataA)
                  With DispDataA
                    If Len(tooltipText) <= 80 Then
                      lstrcpynAAsLong VarPtr(.szText(0)), tooltipText, 80
                    Else
                      .lpszText = StrPtr(tooltipText)
                    End If
                    .hinst = 0
                    If bStore Then .uFlags = .uFlags Or TTF_DI_SETITEM
                  End With
                  CopyMemory lParam, VarPtr(DispDataA), LenB(DispDataA)
                End If
              Case TTN_NEEDTEXTW
                tooltipText = ""
                RaiseEvent NeedText(tooltipText, bStore)
                If tooltipText <> "" Then
                  CopyMemory VarPtr(DispDataW), lParam, LenB(DispDataW)
                  With DispDataW
                    If Len(tooltipText) <= 80 Then
                      lstrcpynWAsLong VarPtr(.szText(0)), StrPtr(tooltipText), 80
                    Else
                      .lpszText = StrPtr(tooltipText)
                    End If
                    .hinst = 0
                    If bStore Then .uFlags = .uFlags Or TTF_DI_SETITEM
                  End With
                  CopyMemory lParam, VarPtr(DispDataW), LenB(DispDataW)
                End If
              Case TTN_POP
                RaiseEvent BeforeHide
            #End If
            Case TTN_SHOW
              #If CustomDraw Then
                If Me.hBitmap Then ResizeToolTipWithBitmap
              #End If

              #If FullFeatured Then
                RaiseEvent BeforeShow
              #End If
          End Select
        End If

      Case WM_NOTIFYFORMAT
        If lParam = NF_QUERY Then
          #If Unicode Then
            ret = NFR_UNICODE
          #Else
            ret = NFR_ANSI
          #End If
          callOldWinMain = False
        End If
    End Select

    If callOldWinMain Then
      ret = CallWindowProc(oldWinMainParent, hWnd, message, wParam, lParam)
    End If

    WinMain_Parent = ret
  End Function

  Private Function WinMain_ToolTip(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    #If CustomDraw Then
      Const WM_PAINT = &HF
    #End If
    Dim callOldWinMain As Boolean
    Dim ret As Long

    callOldWinMain = True
    #If FullFeatured Then
      RaiseEvent ProcessMessage(oldWinMain, hWnd, message, wParam, lParam, callOldWinMain, ret)
    #End If

    If callOldWinMain Then
      #If CustomDraw Then
        Select Case message
          Case WM_PAINT
            If Me.hBitmap Then
              WinMain_ToolTip = HandleWM_PAINT(oldWinMain, hWnd, wParam, lParam)
              Exit Function
            End If
        End Select
      #End If

      ret = CallWindowProc(oldWinMain, hWnd, message, wParam, lParam)
    End If

    WinMain_ToolTip = ret
  End Function
#End If
