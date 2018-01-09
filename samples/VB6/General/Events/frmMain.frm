VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.10#0"; "EditCtlsU.ocx"
Object = "{EA57D88C-8144-415A-9666-B7067B74C295}#1.10#0"; "EditCtlsA.ocx"
Begin VB.Form frmMain 
   Caption         =   "EditControls 1.10 - Events Sample"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'Bildschirmmitte
   Begin EditCtlsLibUCtl.TextBox txtLog 
      Height          =   4815
      Left            =   7320
      TabIndex        =   10
      Top             =   120
      Width           =   4455
      _cx             =   7858
      _cy             =   8493
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   7179
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ScrollBars      =   3
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0000
      Text            =   "frmMain.frx":0020
   End
   Begin EditCtlsLibACtl.UpDownTextBox UpDownTxtBoxA 
      Height          =   300
      Left            =   4560
      TabIndex        =   9
      Top             =   5400
      Width           =   975
      _cx             =   1720
      _cy             =   529
      AcceptNumbersOnly=   0   'False
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutomaticallyCorrectValue=   -1  'True
      AutomaticallySetText=   -1  'True
      AutoScrolling   =   2
      BackColor       =   -2147483643
      Base            =   0
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      CurrentValue    =   0
      DetectDoubleClicks=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      GroupDigits     =   -1  'True
      HAlignment      =   2
      HotTracking     =   0   'False
      HoverTime       =   -1
      IMEMode         =   -1
      LeftMargin      =   -1
      Maximum         =   1000
      MaxTextLength   =   -1
      Minimum         =   0
      Modified        =   0   'False
      MousePointer    =   0
      Orientation     =   1
      ProcessArrowKeys=   -1  'True
      ProcessContextMenuKeys=   -1  'True
      ReadOnlyTextBox =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      UpDownPosition  =   1
      UseSystemFont   =   -1  'True
      WrapAtBoundaries=   0   'False
      CueBanner       =   "frmMain.frx":0040
      Text            =   "frmMain.frx":0060
   End
   Begin EditCtlsLibACtl.HotKeyBox HKBoxA 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   5400
      Width           =   2655
      _cx             =   4683
      _cy             =   503
      Appearance      =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      DefaultModifierKeys=   0
      DetectDoubleClicks=   -1  'True
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverTime       =   -1
      InvalidKeyCombinations=   0
      MousePointer    =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
   End
   Begin EditCtlsLibUCtl.HotKeyBox HKBoxU 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
      _cx             =   4683
      _cy             =   503
      Appearance      =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      DefaultModifierKeys=   0
      DetectDoubleClicks=   -1  'True
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverTime       =   -1
      InvalidKeyCombinations=   0
      MousePointer    =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
   End
   Begin EditCtlsLibACtl.IPAddressBox IPAddressA 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
      _cx             =   2990
      _cy             =   450
      Appearance      =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HoverTime       =   -1
      IMEMode         =   -1
      MousePointer    =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      Address         =   "frmMain.frx":0082
   End
   Begin EditCtlsLibUCtl.IPAddressBox IPAddressU 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
      _cx             =   2990
      _cy             =   450
      Appearance      =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HoverTime       =   -1
      IMEMode         =   -1
      MousePointer    =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      Address         =   "frmMain.frx":00A2
   End
   Begin EditCtlsLibACtl.TextBox TxtBoxA 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   5040
      Width           =   7215
      _cx             =   2646
      _cy             =   503
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
      DisabledEvents  =   0
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   1
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":00C2
      Text            =   "frmMain.frx":011C
   End
   Begin EditCtlsLibACtl.TextBox TxtBoxA 
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   7215
      _cx             =   12726
      _cy             =   4048
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   -1  'True
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   3
      SelectedTextMousePointer=   1
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":013C
      Text            =   "frmMain.frx":015C
   End
   Begin EditCtlsLibUCtl.TextBox TxtBoxU 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   7215
      _cx             =   2646
      _cy             =   503
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
      DisabledEvents  =   0
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   1
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":017C
      Text            =   "frmMain.frx":01D6
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   5400
      Width           =   975
   End
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
      Left            =   8430
      TabIndex        =   12
      Top             =   5280
      Width           =   2415
   End
   Begin EditCtlsLibUCtl.TextBox TxtBoxU 
      Height          =   1575
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _cx             =   12726
      _cy             =   2778
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   -1  'True
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   3
      SelectedTextMousePointer=   1
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":01F6
      Text            =   "frmMain.frx":0216
   End
   Begin EditCtlsLibUCtl.UpDownTextBox UpDownTxtBoxU 
      Height          =   300
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   975
      _cx             =   1720
      _cy             =   529
      AcceptNumbersOnly=   0   'False
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutomaticallyCorrectValue=   -1  'True
      AutomaticallySetText=   -1  'True
      AutoScrolling   =   2
      BackColor       =   -2147483643
      Base            =   0
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      CurrentValue    =   0
      DetectDoubleClicks=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   0
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      GroupDigits     =   -1  'True
      HAlignment      =   2
      HotTracking     =   0   'False
      HoverTime       =   -1
      IMEMode         =   -1
      LeftMargin      =   -1
      Maximum         =   1000
      MaxTextLength   =   -1
      Minimum         =   0
      Modified        =   0   'False
      MousePointer    =   0
      Orientation     =   1
      ProcessArrowKeys=   -1  'True
      ProcessContextMenuKeys=   -1  'True
      ReadOnlyTextBox =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      UpDownPosition  =   1
      UseSystemFont   =   -1  'True
      WrapAtBoundaries=   0   'False
      CueBanner       =   "frmMain.frx":0236
      Text            =   "frmMain.frx":0256
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Const CF_OEMTEXT = 7
  Private Const CF_UNICODETEXT = 13

  Private Const MAX_PATH = 260


  Private bLog As Boolean
  Private objActiveCtl As Object


  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub Form_Click()
  '
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    cx = 0.4 * Me.ScaleWidth
    txtLog.Move Me.ScaleWidth - cx, 0, cx, Me.ScaleHeight - cmdAbout.Height - 10
    cmdAbout.Move txtLog.Left + (cx / 2) - cmdAbout.Width / 2, txtLog.Top + txtLog.Height + 5
    chkLog.Move txtLog.Left, cmdAbout.Top + 5
    TxtBoxU(0).Move 0, 0, txtLog.Left - 5, (Me.ScaleHeight - 5) / 2 - UpDownTxtBoxU.Height - 5 - TxtBoxU(1).Height - 5
    TxtBoxU(1).Move 0, TxtBoxU(0).Top + TxtBoxU(0).Height + 5, TxtBoxU(0).Width
    IPAddressU.Move 0, TxtBoxU(1).Top + TxtBoxU(1).Height + 5
    HKBoxU.Move HKBoxU.Left, TxtBoxU(1).Top + TxtBoxU(1).Height + 3
    UpDownTxtBoxU.Move UpDownTxtBoxU.Left, TxtBoxU(1).Top + TxtBoxU(1).Height + 3
    TxtBoxA(0).Move 0, UpDownTxtBoxU.Top + UpDownTxtBoxU.Height + 5, TxtBoxU(0).Width, TxtBoxU(0).Height
    TxtBoxA(1).Move 0, TxtBoxA(0).Top + TxtBoxA(0).Height + 5, TxtBoxU(0).Width
    IPAddressA.Move 0, TxtBoxA(1).Top + TxtBoxA(1).Height + 5
    HKBoxA.Move HKBoxA.Left, TxtBoxA(1).Top + TxtBoxA(1).Height + 3
    UpDownTxtBoxA.Move UpDownTxtBoxA.Left, TxtBoxA(1).Top + TxtBoxA(1).Height + 3
  End If
End Sub

Private Sub HKBoxA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "HKBoxA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub HKBoxA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "HKBoxA_DragDrop"
End Sub

Private Sub HKBoxA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "HKBoxA_DragOver"
End Sub

Private Sub HKBoxA_GotFocus()
  AddLogEntry "HKBoxA_GotFocus"
End Sub

Private Sub HKBoxA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "HKBoxA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub HKBoxA_KeyPress(keyAscii As Integer)
  AddLogEntry "HKBoxA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub HKBoxA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "HKBoxA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub HKBoxA_LostFocus()
  AddLogEntry "HKBoxA_LostFocus"
End Sub

Private Sub HKBoxA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "HKBoxA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As EditCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "HKBoxA_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub HKBoxA_OLEDragDrop(ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    HKBoxA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub HKBoxA_OLEDragEnter(ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub HKBoxA_OLEDragLeave(ByVal data As EditCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub HKBoxA_OLEDragMouseMove(ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub HKBoxA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "HKBoxA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub HKBoxA_ResizedControlWindow()
  AddLogEntry "HKBoxA_ResizedControlWindow"
End Sub

Private Sub HKBoxA_Validate(Cancel As Boolean)
  AddLogEntry "HKBoxA_Validate"
End Sub

Private Sub HKBoxA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "HKBoxU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub HKBoxU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "HKBoxU_DragDrop"
End Sub

Private Sub HKBoxU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "HKBoxU_DragOver"
End Sub

Private Sub HKBoxU_GotFocus()
  AddLogEntry "HKBoxU_GotFocus"
End Sub

Private Sub HKBoxU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "HKBoxU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub HKBoxU_KeyPress(keyAscii As Integer)
  AddLogEntry "HKBoxU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub HKBoxU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "HKBoxU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub HKBoxU_LostFocus()
  AddLogEntry "HKBoxU_LostFocus"
End Sub

Private Sub HKBoxU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "HKBoxU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As EditCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "HKBoxU_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub HKBoxU_OLEDragDrop(ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    HKBoxU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub HKBoxU_OLEDragEnter(ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub HKBoxU_OLEDragLeave(ByVal data As EditCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub HKBoxU_OLEDragMouseMove(ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "HKBoxU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub HKBoxU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "HKBoxU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub HKBoxU_ResizedControlWindow()
  AddLogEntry "HKBoxU_ResizedControlWindow"
End Sub

Private Sub HKBoxU_Validate(Cancel As Boolean)
  AddLogEntry "HKBoxU_Validate"
End Sub

Private Sub HKBoxU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub HKBoxU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "HKBoxU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_AddressChanged(ByVal editBoxIndex As Long, newFieldValue As Long)
  AddLogEntry "IPAddressA_AddressChanged: editBoxIndex=" & editBoxIndex & ", newFieldValue=" & newFieldValue
End Sub

Private Sub IPAddressA_Click(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_Click: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_ContextMenu(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "IPAddressA_ContextMenu: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub IPAddressA_DblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_DblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "IPAddressA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub IPAddressA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "IPAddressA_DragDrop"
End Sub

Private Sub IPAddressA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "IPAddressA_DragOver"
End Sub

Private Sub IPAddressA_FieldTextChanged(ByVal editBoxIndex As Long)
  AddLogEntry "IPAddressA_FieldTextChanged: editBoxIndex=" & editBoxIndex
End Sub

Private Sub IPAddressA_GotFocus()
  AddLogEntry "IPAddressA_GotFocus"
End Sub

Private Sub IPAddressA_KeyDown(ByVal editBoxIndex As Long, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "IPAddressA_KeyDown: editBoxIndex=" & editBoxIndex & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub IPAddressA_KeyPress(ByVal editBoxIndex As Long, keyAscii As Integer)
  AddLogEntry "IPAddressA_KeyPress: editBoxIndex=" & editBoxIndex & ", keyAscii=" & keyAscii
End Sub

Private Sub IPAddressA_KeyUp(ByVal editBoxIndex As Long, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "IPAddressA_KeyUp: editBoxIndex=" & editBoxIndex & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub IPAddressA_LostFocus()
  AddLogEntry "IPAddressA_LostFocus"
End Sub

Private Sub IPAddressA_MClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_MClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MDblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_MDblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MouseDown(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_MouseDown: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MouseEnter(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_MouseEnter: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MouseHover(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_MouseHover: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MouseLeave(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_MouseLeave: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MouseMove(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "IPAddressA_MouseMove: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MouseUp(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_MouseUp: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_MouseWheel(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As EditCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "IPAddressA_MouseWheel: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub IPAddressA_OLEDragDrop(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressA_OLEDragDrop: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    IPAddressA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub IPAddressA_OLEDragEnter(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressA_OLEDragEnter: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub IPAddressA_OLEDragLeave(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressA_OLEDragLeave: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub IPAddressA_OLEDragMouseMove(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressA_OLEDragMouseMove: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub IPAddressA_RClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_RClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_RDblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_RDblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "IPAddressA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub IPAddressA_ResizedControlWindow()
  AddLogEntry "IPAddressA_ResizedControlWindow"
End Sub

Private Sub IPAddressA_Validate(Cancel As Boolean)
  AddLogEntry "IPAddressA_Validate"
End Sub

Private Sub IPAddressA_XClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_XClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressA_XDblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressA_XDblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_AddressChanged(ByVal editBoxIndex As Long, newFieldValue As Long)
  AddLogEntry "IPAddressU_AddressChanged: editBoxIndex=" & editBoxIndex & ", newFieldValue=" & newFieldValue
End Sub

Private Sub IPAddressU_Click(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_Click: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_ContextMenu(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "IPAddressU_ContextMenu: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub IPAddressU_DblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_DblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "IPAddressU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub IPAddressU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "IPAddressU_DragDrop"
End Sub

Private Sub IPAddressU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "IPAddressU_DragOver"
End Sub

Private Sub IPAddressU_FieldTextChanged(ByVal editBoxIndex As Long)
  AddLogEntry "IPAddressU_FieldTextChanged: editBoxIndex=" & editBoxIndex
End Sub

Private Sub IPAddressU_GotFocus()
  AddLogEntry "IPAddressU_GotFocus"
End Sub

Private Sub IPAddressU_KeyDown(ByVal editBoxIndex As Long, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "IPAddressU_KeyDown: editBoxIndex=" & editBoxIndex & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub IPAddressU_KeyPress(ByVal editBoxIndex As Long, keyAscii As Integer)
  AddLogEntry "IPAddressU_KeyPress: editBoxIndex=" & editBoxIndex & ", keyAscii=" & keyAscii
End Sub

Private Sub IPAddressU_KeyUp(ByVal editBoxIndex As Long, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "IPAddressU_KeyUp: editBoxIndex=" & editBoxIndex & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub IPAddressU_LostFocus()
  AddLogEntry "IPAddressU_LostFocus"
End Sub

Private Sub IPAddressU_MClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_MClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MDblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_MDblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MouseDown(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_MouseDown: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MouseEnter(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_MouseEnter: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MouseHover(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_MouseHover: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MouseLeave(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_MouseLeave: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MouseMove(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "IPAddressU_MouseMove: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MouseUp(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_MouseUp: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_MouseWheel(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As EditCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "IPAddressU_MouseWheel: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub IPAddressU_OLEDragDrop(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressU_OLEDragDrop: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    IPAddressU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub IPAddressU_OLEDragEnter(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressU_OLEDragEnter: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub IPAddressU_OLEDragLeave(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressU_OLEDragLeave: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub IPAddressU_OLEDragMouseMove(ByVal editBoxIndex As Long, ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "IPAddressU_OLEDragMouseMove: editBoxIndex=" & editBoxIndex & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub IPAddressU_RClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_RClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_RDblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_RDblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "IPAddressU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub IPAddressU_ResizedControlWindow()
  AddLogEntry "IPAddressU_ResizedControlWindow"
End Sub

Private Sub IPAddressU_Validate(Cancel As Boolean)
  AddLogEntry "IPAddressU_Validate"
End Sub

Private Sub IPAddressU_XClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_XClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub IPAddressU_XDblClick(ByVal editBoxIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "IPAddressU_XDblClick: editBoxIndex=" & editBoxIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_AbortedDrag(Index As Integer)
  AddLogEntry "TxtBoxA_AbortedDrag: Index=" & Index
End Sub

Private Sub TxtBoxA_BeforeDrawText(Index As Integer)
  AddLogEntry "TxtBoxA_BeforeDrawText: Index=" & Index
End Sub

Private Sub TxtBoxA_BeginDrag(Index As Integer, ByVal firstChar As Long, ByVal lastChar As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim txt As String

  txt = Mid$(TxtBoxA(Index).Text, firstChar + 1, lastChar - firstChar + 1)
  AddLogEntry "TxtBoxA_BeginDrag: Index=" & Index & ", firstChar=" & firstChar & "lastChar=" & lastChar & " (" & txt & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  TxtBoxA(Index).OLEDrag , , , firstChar, lastChar
End Sub

Private Sub TxtBoxA_BeginRDrag(Index As Integer, ByVal firstChar As Long, ByVal lastChar As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim txt As String

  txt = Mid$(TxtBoxA(Index).Text, firstChar + 1, lastChar - firstChar + 1)
  AddLogEntry "TxtBoxA_BeginRDrag: Index=" & Index & ", firstChar=" & firstChar & "lastChar=" & lastChar & " (" & txt & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_Click(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_Click: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_ContextMenu(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "TxtBoxA_ContextMenu: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub TxtBoxA_DblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_DblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_DestroyedControlWindow(Index As Integer, ByVal hWnd As Long)
  AddLogEntry "TxtBoxA_DestroyedControlWindow: Index=" & Index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TxtBoxA_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
  AddLogEntry "TxtBoxA_DragDrop: Index=" & Index
End Sub

Private Sub TxtBoxA_DragMouseMove(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  AddLogEntry "TxtBoxA_DragMouseMove: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
End Sub

Private Sub TxtBoxA_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "TxtBoxA_DragOver: Index=" & Index
End Sub

Private Sub TxtBoxA_Drop(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_Drop: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_GotFocus(Index As Integer)
  AddLogEntry "TxtBoxA_GotFocus: Index=" & Index
  Set objActiveCtl = TxtBoxA(Index)
End Sub

Private Sub TxtBoxA_KeyDown(Index As Integer, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TxtBoxA_KeyDown: Index=" & Index & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TxtBoxA_KeyPress(Index As Integer, keyAscii As Integer)
  AddLogEntry "TxtBoxA_KeyPress: Index=" & Index & ", keyAscii=" & keyAscii
End Sub

Private Sub TxtBoxA_KeyUp(Index As Integer, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TxtBoxA_KeyUp: Index=" & Index & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TxtBoxA_LostFocus(Index As Integer)
  AddLogEntry "TxtBoxA_LostFocus: Index=" & Index
End Sub

Private Sub TxtBoxA_MClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_MClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MDblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_MDblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MouseDown(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_MouseDown: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MouseEnter(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_MouseEnter: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MouseHover(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_MouseHover: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MouseLeave(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_MouseLeave: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MouseMove(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "TxtBoxA_MouseMove: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MouseUp(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_MouseUp: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_MouseWheel(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As EditCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "TxtBoxA_MouseWheel: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub TxtBoxA_OLECompleteDrag(Index As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject, ByVal performedEffect As EditCtlsLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLECompleteDrag: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub TxtBoxA_OLEDragDrop(Index As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLEDragDrop: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    TxtBoxA(Index).FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub TxtBoxA_OLEDragEnter(Index As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLEDragEnter: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub TxtBoxA_OLEDragEnterPotentialTarget(Index As Integer, ByVal hWndPotentialTarget As Long)
  AddLogEntry "TxtBoxA_OLEDragEnterPotentialTarget: Index=" & Index & ", hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub TxtBoxA_OLEDragLeave(Index As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLEDragLeave: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TxtBoxA_OLEDragLeavePotentialTarget(Index As Integer)
  AddLogEntry "TxtBoxA_OLEDragLeavePotentialTarget: Index=" & Index
End Sub

Private Sub TxtBoxA_OLEDragMouseMove(Index As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLEDragMouseMove: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub TxtBoxA_OLEGiveFeedback(Index As Integer, ByVal effect As EditCtlsLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "TxtBoxA_OLEGiveFeedback: Index=" & Index & ", effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub TxtBoxA_OLEQueryContinueDrag(Index As Integer, ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As EditCtlsLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "TxtBoxA_OLEQueryContinueDrag: Index=" & Index & ", pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub TxtBoxA_OLEReceivedNewData(IndexCtl As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLEReceivedNewData: Index=" & IndexCtl & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", Index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub TxtBoxA_OLESetData(IndexCtl As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLESetData: Index=" & IndexCtl & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", Index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub TxtBoxA_OLEStartDrag(Index As Integer, ByVal data As EditCtlsLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "TxtBoxA_OLEStartDrag: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub TxtBoxA_OutOfMemory(Index As Integer)
  AddLogEntry "TxtBoxA_OutOfMemory: Index=" & Index
End Sub

Private Sub TxtBoxA_RClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_RClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_RDblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_RDblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_RecreatedControlWindow(Index As Integer, ByVal hWnd As Long)
  AddLogEntry "TxtBoxA_RecreatedControlWindow: Index=" & Index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TxtBoxA_ResizedControlWindow(Index As Integer)
  AddLogEntry "TxtBoxA_ResizedControlWindow: Index=" & Index
End Sub

Private Sub TxtBoxA_Scrolling(Index As Integer, ByVal axis As EditCtlsLibACtl.ScrollAxisConstants)
  AddLogEntry "TxtBoxA_Scrolling: Index=" & Index & ", axis=" & axis
End Sub

Private Sub TxtBoxA_TextChanged(Index As Integer)
  AddLogEntry "TxtBoxA_TextChanged: Index=" & Index
End Sub

Private Sub TxtBoxA_TruncatedText(Index As Integer)
  AddLogEntry "TxtBoxA_TruncatedText: Index=" & Index
End Sub

Private Sub TxtBoxA_Validate(Index As Integer, Cancel As Boolean)
  AddLogEntry "TxtBoxA_Validate: Index=" & Index
End Sub

Private Sub TxtBoxA_WritingDirectionChanged(Index As Integer, ByVal newWritingDirection As EditCtlsLibACtl.WritingDirectionConstants)
  AddLogEntry "TxtBoxA_WritingDirectionChanged: Index=" & Index & ", newWritingDirection=" & newWritingDirection
End Sub

Private Sub TxtBoxA_XClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_XClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxA_XDblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxA_XDblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_AbortedDrag(Index As Integer)
  AddLogEntry "TxtBoxU_AbortedDrag: Index=" & Index
End Sub

Private Sub TxtBoxU_BeforeDrawText(Index As Integer)
  AddLogEntry "TxtBoxU_BeforeDrawText: Index=" & Index
End Sub

Private Sub TxtBoxU_BeginDrag(Index As Integer, ByVal firstChar As Long, ByVal lastChar As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim txt As String

  txt = Mid$(TxtBoxU(Index).Text, firstChar + 1, lastChar - firstChar + 1)
  AddLogEntry "TxtBoxU_BeginDrag: Index=" & Index & ", firstChar=" & firstChar & ", lastChar=" & lastChar & " (" & txt & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  TxtBoxU(Index).OLEDrag , , , firstChar, lastChar
End Sub

Private Sub TxtBoxU_BeginRDrag(Index As Integer, ByVal firstChar As Long, ByVal lastChar As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim txt As String

  txt = Mid$(TxtBoxU(Index).Text, firstChar + 1, lastChar - firstChar + 1)
  AddLogEntry "TxtBoxU_BeginRDrag: Index=" & Index & ", firstChar=" & firstChar & ", lastChar=" & lastChar & " (" & txt & "), button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_Click(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_Click: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_ContextMenu(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "TxtBoxU_ContextMenu: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub TxtBoxU_DblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_DblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_DestroyedControlWindow(Index As Integer, ByVal hWnd As Long)
  AddLogEntry "TxtBoxU_DestroyedControlWindow: Index=" & Index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TxtBoxU_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
  AddLogEntry "TxtBoxU_DragDrop: Index=" & Index
End Sub

Private Sub TxtBoxU_DragMouseMove(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  AddLogEntry "TxtBoxU_DragMouseMove: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
End Sub

Private Sub TxtBoxU_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "TxtBoxU_DragOver: Index=" & Index
End Sub

Private Sub TxtBoxU_Drop(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_Drop: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_GotFocus(Index As Integer)
  AddLogEntry "TxtBoxU_GotFocus: Index=" & Index
  Set objActiveCtl = TxtBoxU(Index)
End Sub

Private Sub TxtBoxU_KeyDown(Index As Integer, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TxtBoxU_KeyDown: Index=" & Index & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TxtBoxU_KeyPress(Index As Integer, keyAscii As Integer)
  AddLogEntry "TxtBoxU_KeyPress: Index=" & Index & ", keyAscii=" & keyAscii
End Sub

Private Sub TxtBoxU_KeyUp(Index As Integer, keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TxtBoxU_KeyUp: Index=" & Index & ", keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TxtBoxU_LostFocus(Index As Integer)
  AddLogEntry "TxtBoxU_LostFocus: Index=" & Index
End Sub

Private Sub TxtBoxU_MClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_MClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MDblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_MDblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MouseDown(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_MouseDown: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MouseEnter(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_MouseEnter: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MouseHover(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_MouseHover: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MouseLeave(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_MouseLeave: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MouseMove(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "TxtBoxU_MouseMove: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MouseUp(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_MouseUp: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_MouseWheel(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As EditCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "TxtBoxU_MouseWheel: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub TxtBoxU_OLECompleteDrag(Index As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject, ByVal performedEffect As EditCtlsLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "TxtBoxU_OLECompleteDrag: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub TxtBoxU_OLEDragDrop(Index As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TxtBoxU_OLEDragDrop: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    TxtBoxU(Index).FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub TxtBoxU_OLEDragEnter(Index As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TxtBoxU_OLEDragEnter: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub TxtBoxU_OLEDragEnterPotentialTarget(Index As Integer, ByVal hWndPotentialTarget As Long)
  AddLogEntry "TxtBoxU_OLEDragEnterPotentialTarget: Index=" & Index & ", hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub TxtBoxU_OLEDragLeave(Index As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "TxtBoxU_OLEDragLeave: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub TxtBoxU_OLEDragLeavePotentialTarget(Index As Integer)
  AddLogEntry "TxtBoxU_OLEDragLeavePotentialTarget: Index=" & Index
End Sub

Private Sub TxtBoxU_OLEDragMouseMove(Index As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TxtBoxU_OLEDragMouseMove: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub TxtBoxU_OLEGiveFeedback(Index As Integer, ByVal effect As EditCtlsLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "TxtBoxU_OLEGiveFeedback: Index=" & Index & ", effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub TxtBoxU_OLEQueryContinueDrag(Index As Integer, ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As EditCtlsLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "TxtBoxU_OLEQueryContinueDrag: Index=" & Index & ", pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub TxtBoxU_OLEReceivedNewData(IndexCtl As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "TxtBoxU_OLEReceivedNewData: Index=" & IndexCtl & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", Index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub TxtBoxU_OLESetData(IndexCtl As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim firstChar As Long
  Dim lastChar As Long
  Dim str As String

  str = "TxtBoxU_OLESetData: Index=" & IndexCtl & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", Index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str

  Select Case formatID
    Case vbCFText, CF_OEMTEXT, CF_UNICODETEXT
      TxtBoxU(IndexCtl).GetDraggedTextRange firstChar, lastChar
      data.SetData formatID, Mid$(TxtBoxU(IndexCtl).Text, firstChar + 1, lastChar - firstChar + 1)
  End Select
End Sub

Private Sub TxtBoxU_OLEStartDrag(Index As Integer, ByVal data As EditCtlsLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "TxtBoxU_OLEStartDrag: Index=" & Index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str

  data.SetData vbCFText
  data.SetData CF_OEMTEXT
  data.SetData CF_UNICODETEXT
End Sub

Private Sub TxtBoxU_OutOfMemory(Index As Integer)
  AddLogEntry "TxtBoxU_OutOfMemory: Index=" & Index
End Sub

Private Sub TxtBoxU_RClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_RClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_RDblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_RDblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_RecreatedControlWindow(Index As Integer, ByVal hWnd As Long)
  AddLogEntry "TxtBoxU_RecreatedControlWindow: Index=" & Index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TxtBoxU_ResizedControlWindow(Index As Integer)
  AddLogEntry "TxtBoxU_ResizedControlWindow: Index=" & Index
End Sub

Private Sub TxtBoxU_Scrolling(Index As Integer, ByVal axis As EditCtlsLibUCtl.ScrollAxisConstants)
  AddLogEntry "TxtBoxU_Scrolling: Index=" & Index & ", axis=" & axis
End Sub

Private Sub TxtBoxU_TextChanged(Index As Integer)
  AddLogEntry "TxtBoxU_TextChanged: Index=" & Index
End Sub

Private Sub TxtBoxU_TruncatedText(Index As Integer)
  AddLogEntry "TxtBoxU_TruncatedText: Index=" & Index
End Sub

Private Sub TxtBoxU_Validate(Index As Integer, Cancel As Boolean)
  AddLogEntry "TxtBoxU_Validate: Index=" & Index
End Sub

Private Sub TxtBoxU_WritingDirectionChanged(Index As Integer, ByVal newWritingDirection As EditCtlsLibUCtl.WritingDirectionConstants)
  AddLogEntry "TxtBoxU_WritingDirectionChanged: Index=" & Index & ", newWritingDirection=" & newWritingDirection
End Sub

Private Sub TxtBoxU_XClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_XClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub TxtBoxU_XDblClick(Index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "TxtBoxU_XDblClick: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub UpDownTxtBoxA_BeforeDrawText()
  AddLogEntry "UpDownTxtBoxA_BeforeDrawText"
End Sub

Private Sub UpDownTxtBoxA_ChangedAccelerators(ByVal Accelerators As EditCtlsLibACtl.IUpDownAccelerators)
  If Accelerators Is Nothing Then
    AddLogEntry "UpDownTxtBoxA_ChangedAccelerators: Accelerators=Nothing"
  Else
    AddLogEntry "UpDownTxtBoxA_ChangedAccelerators: Accelerators=" & Accelerators.Count
  End If
End Sub

Private Sub UpDownTxtBoxA_ChangingAccelerators(ByVal Accelerators As EditCtlsLibACtl.IVirtualUpDownAccelerators, cancelChanges As Boolean)
  If Accelerators Is Nothing Then
    AddLogEntry "UpDownTxtBoxA_ChangingAccelerators: Accelerators=Nothing, cancelChanges=" & cancelChanges
  Else
    AddLogEntry "UpDownTxtBoxA_ChangingAccelerators: Accelerators=" & Accelerators.Count & ", cancelChanges=" & cancelChanges
  End If
End Sub

Private Sub UpDownTxtBoxA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "UpDownTxtBoxA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub UpDownTxtBoxA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "UpDownTxtBoxA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub UpDownTxtBoxA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "UpDownTxtBoxA_DragDrop"
End Sub

Private Sub UpDownTxtBoxA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "UpDownTxtBoxA_DragOver"
End Sub

Private Sub UpDownTxtBoxA_GotFocus()
  AddLogEntry "UpDownTxtBoxA_GotFocus"
End Sub

Private Sub UpDownTxtBoxA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "UpDownTxtBoxA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub UpDownTxtBoxA_KeyPress(keyAscii As Integer)
  AddLogEntry "UpDownTxtBoxA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub UpDownTxtBoxA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "UpDownTxtBoxA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub UpDownTxtBoxA_LostFocus()
  AddLogEntry "UpDownTxtBoxA_LostFocus"
End Sub

Private Sub UpDownTxtBoxA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  'AddLogEntry "UpDownTxtBoxA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants, ByVal scrollAxis As EditCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "UpDownTxtBoxA_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub UpDownTxtBoxA_OLEDragDrop(ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    UpDownTxtBoxA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub UpDownTxtBoxA_OLEDragEnter(ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub UpDownTxtBoxA_OLEDragLeave(ByVal data As EditCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub UpDownTxtBoxA_OLEDragMouseMove(ByVal data As EditCtlsLibACtl.IOLEDataObject, effect As EditCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub UpDownTxtBoxA_OutOfMemory()
  AddLogEntry "UpDownTxtBoxA_OutOfMemory"
End Sub

Private Sub UpDownTxtBoxA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "UpDownTxtBoxA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub UpDownTxtBoxA_ResizedControlWindow()
  AddLogEntry "UpDownTxtBoxA_ResizedControlWindow"
End Sub

Private Sub UpDownTxtBoxA_TextChanged()
  AddLogEntry "UpDownTxtBoxA_TextChanged"
End Sub

Private Sub UpDownTxtBoxA_TruncatedText()
  AddLogEntry "UpDownTxtBoxA_TruncatedText"
End Sub

Private Sub UpDownTxtBoxA_Validate(Cancel As Boolean)
  AddLogEntry "UpDownTxtBoxA_Validate"
End Sub

Private Sub UpDownTxtBoxA_ValueChanged()
  AddLogEntry "UpDownTxtBoxA_ValueChanged"
End Sub

Private Sub UpDownTxtBoxA_ValueChanging(ByVal CurrentValue As Long, ByVal delta As Long, cancelChange As Boolean)
  AddLogEntry "UpDownTxtBoxA_ValueChanging: currentValue=" & CurrentValue & ", delta=" & delta & ", cancelChange=" & cancelChange
End Sub

Private Sub UpDownTxtBoxA_WritingDirectionChanged(ByVal newWritingDirection As EditCtlsLibACtl.WritingDirectionConstants)
  AddLogEntry "UpDownTxtBoxA_WritingDirectionChanged: newWritingDirection=" & newWritingDirection
End Sub

Private Sub UpDownTxtBoxA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibACtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_BeforeDrawText()
  AddLogEntry "UpDownTxtBoxU_BeforeDrawText"
End Sub

Private Sub UpDownTxtBoxU_ChangedAccelerators(ByVal Accelerators As EditCtlsLibUCtl.IUpDownAccelerators)
  If Accelerators Is Nothing Then
    AddLogEntry "UpDownTxtBoxU_ChangedAccelerators: Accelerators=Nothing"
  Else
    AddLogEntry "UpDownTxtBoxU_ChangedAccelerators: Accelerators=" & Accelerators.Count
  End If
End Sub

Private Sub UpDownTxtBoxU_ChangingAccelerators(ByVal Accelerators As EditCtlsLibUCtl.IVirtualUpDownAccelerators, cancelChanges As Boolean)
  If Accelerators Is Nothing Then
    AddLogEntry "UpDownTxtBoxU_ChangingAccelerators: Accelerators=Nothing, cancelChanges=" & cancelChanges
  Else
    AddLogEntry "UpDownTxtBoxU_ChangingAccelerators: Accelerators=" & Accelerators.Count & ", cancelChanges=" & cancelChanges
  End If
End Sub

Private Sub UpDownTxtBoxU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "UpDownTxtBoxU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub UpDownTxtBoxU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "UpDownTxtBoxU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub UpDownTxtBoxU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "UpDownTxtBoxU_DragDrop"
End Sub

Private Sub UpDownTxtBoxU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "UpDownTxtBoxU_DragOver"
End Sub

Private Sub UpDownTxtBoxU_GotFocus()
  AddLogEntry "UpDownTxtBoxU_GotFocus"
End Sub

Private Sub UpDownTxtBoxU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "UpDownTxtBoxU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub UpDownTxtBoxU_KeyPress(keyAscii As Integer)
  AddLogEntry "UpDownTxtBoxU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub UpDownTxtBoxU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "UpDownTxtBoxU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub UpDownTxtBoxU_LostFocus()
  AddLogEntry "UpDownTxtBoxU_LostFocus"
End Sub

Private Sub UpDownTxtBoxU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  'AddLogEntry "UpDownTxtBoxU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants, ByVal scrollAxis As EditCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "UpDownTxtBoxU_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub UpDownTxtBoxU_OLEDragDrop(ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    UpDownTxtBoxU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub UpDownTxtBoxU_OLEDragEnter(ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub UpDownTxtBoxU_OLEDragLeave(ByVal data As EditCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub UpDownTxtBoxU_OLEDragMouseMove(ByVal data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "UpDownTxtBoxU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub UpDownTxtBoxU_OutOfMemory()
  AddLogEntry "UpDownTxtBoxU_OutOfMemory"
End Sub

Private Sub UpDownTxtBoxU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "UpDownTxtBoxU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub UpDownTxtBoxU_ResizedControlWindow()
  AddLogEntry "UpDownTxtBoxU_ResizedControlWindow"
End Sub

Private Sub UpDownTxtBoxU_TextChanged()
  AddLogEntry "UpDownTxtBoxU_TextChanged"
End Sub

Private Sub UpDownTxtBoxU_TruncatedText()
  AddLogEntry "UpDownTxtBoxU_TruncatedText"
End Sub

Private Sub UpDownTxtBoxU_Validate(Cancel As Boolean)
  AddLogEntry "UpDownTxtBoxU_Validate"
End Sub

Private Sub UpDownTxtBoxU_ValueChanged()
  AddLogEntry "UpDownTxtBoxU_ValueChanged"
End Sub

Private Sub UpDownTxtBoxU_ValueChanging(ByVal CurrentValue As Long, ByVal delta As Long, cancelChange As Boolean)
  AddLogEntry "UpDownTxtBoxU_ValueChanging: currentValue=" & CurrentValue & ", delta=" & delta & ", cancelChange=" & cancelChange
End Sub

Private Sub UpDownTxtBoxU_WritingDirectionChanged(ByVal newWritingDirection As EditCtlsLibUCtl.WritingDirectionConstants)
  AddLogEntry "UpDownTxtBoxU_WritingDirectionChanged: newWritingDirection=" & newWritingDirection
End Sub

Private Sub UpDownTxtBoxU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub UpDownTxtBoxU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As EditCtlsLibUCtl.HitTestConstants)
  AddLogEntry "UpDownTxtBoxU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long

  If bLog Then
    txtLog.DontRedraw = True
    If txtLog.GetLineCount > 50 Then
      pos = txtLog.GetFirstCharOfLine(txtLog.GetLineCount - 50)
      txtLog.SetSelection 0, pos
      txtLog.ReplaceSelectedText ""
    End If
    txtLog.DontRedraw = False
    txtLog.AppendText txt & vbNewLine, True, True
  End If
End Sub
