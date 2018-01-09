VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.10#0"; "EditCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "EditControls 1.10 - Performance Demonstration Sample"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   2  'Bildschirmmitte
   Begin EditCtlsLibUCtl.UpDownTextBox udLines 
      Height          =   255
      Left            =   675
      TabIndex        =   4
      Top             =   4740
      Width           =   975
      _cx             =   1720
      _cy             =   450
      AcceptNumbersOnly=   -1  'True
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
      CueBanner       =   ""
      CurrentValue    =   1000
      DisabledBackColor=   -1
      DisabledEvents  =   13323
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
      Maximum         =   2000
      MaxTextLength   =   -1
      Minimum         =   1
      Modified        =   0   'False
      MousePointer    =   0
      Orientation     =   1
      ProcessArrowKeys=   -1  'True
      ProcessContextMenuKeys=   -1  'True
      ReadOnlyTextBox =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      Text            =   "1.000"
      UpDownPosition  =   1
      UseSystemFont   =   -1  'True
      WrapAtBoundaries=   0   'False
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "&Fill TextBoxes"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox VBTxtBox 
      BackColor       =   &H8000000F&
      Height          =   3855
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      ToolTipText     =   "VB6 TextBox"
      Top             =   600
      Width           =   3975
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
      Left            =   7320
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin EditCtlsLibUCtl.TextBox TxtBoxU 
      Height          =   3885
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "EditControls TextBox"
      Top             =   600
      Width           =   4695
      _cx             =   8281
      _cy             =   6853
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   0   'False
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      CueBanner       =   ""
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
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   3
      SelectedTextMousePointer=   1
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      Text            =   ""
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
   End
   Begin VB.Label lblDescr 
      BackStyle       =   0  'Transparent
      Caption         =   "Task: Insert x separate lines as fast as possible while keeping the cursor at the end of the text."
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Lines:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   420
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill time: 0 ms/0 ms"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   4800
      Width           =   1350
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub cmdAbout_Click()
  TxtBoxU.About
End Sub

Private Sub cmdFill_Click()
  Dim cLines As Long
  Dim eEditCtls As Long
  Dim eVB6 As Long
  Dim iLine As Long
  Dim sEditCtls As Long
  Dim sVB6 As Long

  cmdFill.Enabled = False
  TxtBoxU.Text = ""
  VBTxtBox.Text = ""
  cLines = udLines.CurrentValue
  DoEvents

  sEditCtls = GetTickCount
  For iLine = 1 To cLines
    TxtBoxU.AppendText "This is line " & CStr(iLine) & vbNewLine, True, True
  Next iLine
  eEditCtls = GetTickCount

  DoEvents
  DoEvents
  DoEvents

  sVB6 = GetTickCount
  For iLine = 1 To cLines
    VBTxtBox.Text = VBTxtBox.Text & "This is line " & CStr(iLine) & vbNewLine
    VBTxtBox.SelStart = 65535
  Next iLine
  eVB6 = GetTickCount

  lblTime.Caption = "Fill time: " & CStr(eEditCtls - sEditCtls) & " ms/" & CStr(eVB6 - sVB6) & " ms"
  cmdFill.Enabled = True
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    cmdAbout.Move Me.ScaleWidth - cmdAbout.Width - 5, Me.ScaleHeight - cmdAbout.Height - 5
    lblDescr(1).Move 5, Me.ScaleHeight - lblDescr(1).Height - 10
    udLines.Move lblDescr(1).Left + lblDescr(1).Width + 3, lblDescr(1).Top - 4
    cmdFill.Move udLines.Left + udLines.Width + 5, cmdAbout.Top
    lblTime.Move cmdFill.Left + cmdFill.Width + 10, lblDescr(1).Top
    TxtBoxU.Move 0, TxtBoxU.Top, Me.ScaleWidth / 2 - 3, cmdAbout.Top - 5 - TxtBoxU.Top
    VBTxtBox.Move TxtBoxU.Left + TxtBoxU.Width + 6, TxtBoxU.Top, TxtBoxU.Width, TxtBoxU.Height
  End If
End Sub
