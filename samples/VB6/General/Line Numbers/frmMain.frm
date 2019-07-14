VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.11#0"; "EditCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "EditControls 1.11 - Line Numbers Sample"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
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
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
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
      Left            =   4920
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin EditCtlsLibUCtl.TextBox TxtBoxU 
      Align           =   3  'Links ausrichten
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _cx             =   8281
      _cy             =   7011
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
      DisabledBackColor=   -1
      DisabledEvents  =   3074
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

  Implements ISubclassedWindow


  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private lineNumbersWidth As Long


  Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function ValidateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMainTxtBoxU
      lRet = HandleMessage_TxtBoxU(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_TxtBoxU(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_PAINT = &HF
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_PAINT
      DrawLineNumbers
  End Select

StdHandler_Ende:
  HandleMessage_TxtBoxU = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_TxtBoxU: ", Err.number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  TxtBoxU.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  #If UseSubClassing Then
    If Not SubclassWindow(TxtBoxU.hWnd, Me, EnumSubclassID.escidFrmMainTxtBoxU) Then
      Debug.Print "Subclassing failed!"
    End If
  #End If

  LoadTextU
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(TxtBoxU.hWnd, EnumSubclassID.escidFrmMainTxtBoxU) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub TxtBoxU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  ' redraw in case the selection has changed
  DrawLineNumbers
End Sub

Private Sub TxtBoxU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  If ScaleX(x, vbTwips, vbPixels) < lineNumbersWidth Then
    TxtBoxU.MousePointer = mpArrow
  Else
    TxtBoxU.MousePointer = mpDefault
  End If

  ' redraw in case the selection has changed
  DrawLineNumbers
End Sub

Private Sub TxtBoxU_RecreatedControlWindow(ByVal hWnd As Long)
  LoadTextU
End Sub

Private Sub TxtBoxU_TextChanged()
  Dim cx As Long

  cx = CalculateLineNumbersWidth()
  If cx <> lineNumbersWidth Then
    lineNumbersWidth = cx
    TxtBoxU.LeftMargin = cx
  Else
    DrawLineNumbers
  End If
End Sub


Private Function CalculateLineNumbersWidth()
  Const DT_CALCRECT = &H400
  Const DT_LEFT = &H0
  Const WM_GETFONT = &H31
  Dim hBMP As Long
  Dim hDC As Long
  Dim hDCMem As Long
  Dim hFont As Long
  Dim hOldBMP As Long
  Dim hOldFont As Long
  Dim number As String
  Dim rc As RECT

  hDC = GetDC(TxtBoxU.hWnd)
  If hDC Then
    rc.Right = 200
    rc.Bottom = 200

    hDCMem = CreateCompatibleDC(hDC)
    If hDCMem Then
      hBMP = CreateCompatibleBitmap(hDC, rc.Right - rc.Left, rc.Bottom - rc.Top)
      If hBMP Then
        hOldBMP = SelectObject(hDCMem, hBMP)
        hFont = SendMessageAsLong(TxtBoxU.hWnd, WM_GETFONT, 0, 0)
        hOldFont = SelectObject(hDCMem, hFont)

        number = CStr(TxtBoxU.GetLineCount())
        DrawText hDCMem, StrPtr(number), -1, rc, DT_LEFT Or DT_CALCRECT
        CalculateLineNumbersWidth = rc.Right - rc.Left + 10

        SelectObject hDCMem, hOldFont
        SelectObject hDCMem, hOldBMP
        DeleteObject hBMP
      End If
      DeleteDC hDCMem
    End If
    ReleaseDC TxtBoxU.hWnd, hDC
  End If
End Function

Private Sub DrawLineNumbers()
  Const COLOR_3DFACE = 15
  Const DT_RIGHT = &H2
  Const SRCCOPY = &HCC0020
  Const Transparent = 1
  Const WM_GETFONT = &H31
  Dim hBMP As Long
  Dim hDC As Long
  Dim hDCMem As Long
  Dim hFont As Long
  Dim hOldBMP As Long
  Dim hOldFont As Long
  Dim i As Long
  Dim numbers As String
  Dim oldBkMode As Long
  Dim oldColor As Long
  Dim rc As RECT
  Dim rc2 As RECT

  hDC = GetDC(TxtBoxU.hWnd)
  If hDC Then
    For i = TxtBoxU.FirstVisibleLine + 1 To TxtBoxU.LastVisibleLine + 1
      numbers = numbers & CStr(i) & vbNewLine
    Next i

    rc.Right = lineNumbersWidth - 2
    rc.Bottom = rc.Top + TxtBoxU.Height

    hDCMem = CreateCompatibleDC(hDC)
    If hDCMem Then
      hBMP = CreateCompatibleBitmap(hDC, rc.Right - rc.Left, rc.Bottom - rc.Top)
      If hBMP Then
        hOldBMP = SelectObject(hDCMem, hBMP)
        oldBkMode = SetBkMode(hDCMem, Transparent)
        hFont = SendMessageAsLong(TxtBoxU.hWnd, WM_GETFONT, 0, 0)
        hOldFont = SelectObject(hDCMem, hFont)
        oldColor = SetTextColor(hDCMem, vbRed)

        FillRect hDCMem, rc, GetSysColorBrush(COLOR_3DFACE)
        rc2 = rc
        rc2.Left = 1
        If Len(TxtBoxU.Text) > 0 Then
          TxtBoxU.CharIndexToPosition TxtBoxU.GetFirstCharOfLine(TxtBoxU.FirstVisibleLine), , rc2.Top
        Else
          rc2.Top = 2
        End If
        rc2.Right = rc2.Right - 5
        DrawText hDCMem, StrPtr(numbers), -1, rc2, DT_RIGHT

        BitBlt hDC, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, hDCMem, 0, 0, SRCCOPY
        ValidateRect TxtBoxU.hWnd, rc

        SetTextColor hDCMem, oldColor
        SelectObject hDCMem, hOldFont
        SetBkMode hDCMem, oldBkMode
        SelectObject hDCMem, hOldBMP
        DeleteObject hBMP
      End If
      DeleteDC hDCMem
    End If
    ReleaseDC TxtBoxU.hWnd, hDC
  End If
End Sub

Private Sub LoadTextU()
  Dim filePath As String
  Dim fn As Long

  If Right$(App.Path, 3) = "bin" Then
    filePath = App.Path & "\clsToolTip.cls"
  Else
    filePath = App.Path & "\bin\clsToolTip.cls"
  End If
  fn = FreeFile
  On Error GoTo FileNotFound
  Open filePath For Input As #fn
    TxtBoxU.Text = Input(LOF(fn), fn)
  Close #fn
  Exit Sub

FileNotFound:
End Sub
