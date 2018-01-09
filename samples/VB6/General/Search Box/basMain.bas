Attribute VB_Name = "basMain"
Option Explicit

  ' This sample was heavily inspired by the "Explorer breadcrumbs in Vista" sample by
  ' Bjarke Viksoe which can be found on http://www.viksoe.dk/code/breadcrumbs.htm


  Public Const S_OK As Long = &H0


  Public Type MARGINS
    cxLeftWidth As Long
    cxRightWidth As Long
    cyTopHeight As Long
    cyBottomHeight As Long
  End Type

  Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Public Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(0 To 31) As Byte
  End Type


  Public Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
  Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
  Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Public Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef pfEnabled As Long) As Long
  Public Declare Function EndPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
  Public Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lprc As RECT, ByVal hbr As Long) As Long
  Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal fnObject As Long) As Long
  Public Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
  Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
  Public Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long


Public Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

Public Function IsCompositionEnabled() As Boolean
  Dim compositionEnabled As Long

  ' retrieves whether Aero Glass is enabled

  If DwmIsCompositionEnabled(compositionEnabled) < 0 Then
    compositionEnabled = 0
  End If
  IsCompositionEnabled = (compositionEnabled <> 0)
End Function

Public Function LoWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value), LenB(ret)

  LoWord = ret
End Function
