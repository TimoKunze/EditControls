VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.11#0"; "EditCtlsU.ocx"
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.0#0"; "TBarCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "EditControls 1.11 - Search Box Sample"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin SearchBox.UCSearchBox SearchBand 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
   End
   Begin TBarCtlsLibUCtl.ReBar NavBar 
      Align           =   1  'Oben ausrichten
      Height          =   660
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5415
      _cx             =   9551
      _cy             =   1164
      AllowBandReordering=   -1  'True
      Appearance      =   0
      AutoUpdateLayout=   -1  'True
      BackColor       =   -1
      BorderStyle     =   0
      DisabledEvents  =   524779
      DisplayBandSeparators=   0   'False
      DisplaySplitter =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      FixedBandHeight =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -1
      HighlightColor  =   -1
      HoverTime       =   -1
      MousePointer    =   0
      Orientation     =   0
      ReplaceMDIFrameMenu=   0
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      SupportOLEDragImages=   -1  'True
      ToggleOnDoubleClick=   0   'False
      UseSystemFont   =   -1  'True
      VerticalSizingGripsOnVerticalOrientation=   -1  'True
   End
   Begin EditCtlsLibUCtl.TextBox Text1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5175
      _cx             =   9128
      _cy             =   2566
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   3075
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   -1  'True
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0000
      Text            =   "frmMain.frx":0020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  ' This sample was heavily inspired by the "Explorer breadcrumbs in Vista" sample by
  ' Bjarke Viksoe which can be found on http://www.viksoe.dk/code/breadcrumbs.htm


  Implements ISubclassedWindow


  Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 256) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
  End Type


  Private Const CX_SEARCH As Long = 200
  Private Const CY_SEARCH As Long = 24
  Private Const CY_NAVBAR As Long = 27


  Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, ByRef pMarInset As MARGINS) As Long
  Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
  Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInfo As OSVERSIONINFOEX) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_ERASEBKGND As Long = &H14
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_ERASEBKGND
      If Not IsCompositionEnabled Then DoPaint wParam
      lRet = 1
      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub Form_Activate()
  Dim windowMargins As MARGINS

  If IsCompositionEnabled Then
    ' extend Aero Glass border into the client area so that the search box sits on the glass area
    windowMargins.cyTopHeight = CY_NAVBAR
    DwmExtendFrameIntoClientArea Me.hWnd, windowMargins
  End If
  Me.Refresh
End Sub

Private Sub Form_Initialize()
  Dim hTheme As Long
  Dim OSVerData As OSVERSIONINFOEX

  InitCommonControls

  With OSVerData
    .dwOSVersionInfoSize = LenB(OSVerData)
    GetVersionEx OSVerData
    If .dwMajorVersion < 6 Then
      MsgBox "This sample requires Windows Vista or newer.", VbMsgBoxStyle.vbCritical, "Error"
      End
    End If
  End With
  hTheme = OpenThemeData(Me.hWnd, StrPtr("BreadcrumbBarComposited::BreadcrumbBar"))
  If hTheme = 0 Then
    MsgBox "Cannot open theme class data." & vbNewLine & "Wrong version of Windows or no theming supported.", VbMsgBoxStyle.vbCritical, "Error"
    End
  End If
  CloseThemeData hTheme
End Sub

Private Sub Form_Load()
  Dim band As ReBarBand

  ' change theme of the rebar to a special one
  SetWindowTheme NavBar.hWnd, StrPtr("NavbarComposited"), 0
  ' subclass the main form in order to draw the background
  If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
    Debug.Print "Subclassing failed!"
  End If

  Set band = NavBar.Bands.Add(, SearchBand.hWnd, , sgvNever, False, , , CX_SEARCH, CX_SEARCH, CY_SEARCH, CY_SEARCH)
End Sub

Private Sub Form_Resize()
  NavBar.Height = CY_NAVBAR
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub SearchBand_ApplyFilter()
  Text1.Text = "Search for: " & SearchBand.Filter
End Sub

Private Sub SearchBand_ClearFilter()
  Text1.Text = "Filter cleared"
End Sub

Private Sub SearchBand_FilterChange()
  Text1.Text = "Filter: " & SearchBand.Filter
End Sub

Private Sub cmdAbout_Click()
  Text1.About
End Sub


Private Sub DoPaint(ByVal hDC As Long)
  Const BLACK_BRUSH As Long = 4
  Const COLOR_BTNFACE As Long = 15
  Const COLOR_GRADIENTACTIVECAPTION As Long = 27
  Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
  Dim rcClient As RECT
  Dim hBrush As Long

  ' at first draw the background of the rebar
  GetClientRect Me.hWnd, rcClient
  rcClient.Bottom = rcClient.Top + CY_NAVBAR
  If IsCompositionEnabled Then
    ' We extended the Glass border into the client area. To make this work entirely, we have to
    ' paint this part of the client area black.
    hBrush = GetStockObject(BLACK_BRUSH)
  Else
    ' draw the rebar's background in a color that looks a bit like Aero Glass
    If GetActiveWindow = Me.hWnd Then
      hBrush = GetSysColorBrush(COLOR_GRADIENTACTIVECAPTION)
    Else
      hBrush = GetSysColorBrush(COLOR_GRADIENTINACTIVECAPTION)
    End If
  End If
  If hBrush Then
    ' do the actual drawing
    FillRect hDC, rcClient, hBrush
  End If
End Sub
