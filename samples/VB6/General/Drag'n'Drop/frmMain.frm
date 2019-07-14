VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.11#0"; "EditCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "EditControls 1.11 - Drag'n'Drop Sample"
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5160
      Top             =   720
   End
   Begin VB.CheckBox chkOLEDragDrop 
      Caption         =   "&OLE Drag'n'Drop"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
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
      AllowDragDrop   =   -1  'True
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
      DisabledEvents  =   7170
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
      Text            =   ""
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   -1  'True
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Text"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move Text"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "C&ancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Enum ChosenMenuItemConstants
    ciCopy
    ciMove
    ciCancel
  End Enum


  Private Const CF_DIBV5 = 17
  Private Const CF_OEMTEXT = 7
  Private Const CF_UNICODETEXT = 13

  Private Const strCLSID_RecycleBin = "{645FF040-5081-101B-9F08-00AA002F954E}"


  Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
  End Type


  Private CFSTR_HTML As Long
  Private CFSTR_HTML2 As Long
  Private CFSTR_LOGICALPERFORMEDDROPEFFECT As Long
  Private CFSTR_MOZILLAHTMLCONTEXT As Long
  Private CFSTR_PERFORMEDDROPEFFECT As Long
  Private CFSTR_TARGETCLSID As Long
  Private CFSTR_TEXTHTML As Long

  Private bRightDrag As Boolean
  Private chosenMenuItem As ChosenMenuItemConstants
  Private CLSID_RecycleBin As GUID
  Private suppressDefaultContextMenu As Boolean
  Private ToolTip As clsToolTip


  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As GUID) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long


Private Sub cmdAbout_Click()
  TxtBoxU.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls

  CFSTR_HTML = RegisterClipboardFormat(StrPtr("HTML Format"))
  CFSTR_HTML2 = RegisterClipboardFormat(StrPtr("HTML (HyperText Markup Language)"))
  CFSTR_LOGICALPERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Logical Performed DropEffect"))
  CFSTR_MOZILLAHTMLCONTEXT = RegisterClipboardFormat(StrPtr("text/_moz_htmlcontext"))
  CFSTR_PERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Performed DropEffect"))
  CFSTR_TARGETCLSID = RegisterClipboardFormat(StrPtr("TargetCLSID"))
  CFSTR_TEXTHTML = RegisterClipboardFormat(StrPtr("text/html"))

  CLSIDFromString StrPtr(strCLSID_RecycleBin), CLSID_RecycleBin
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
  Dim firstChar As Long
  Dim lastChar As Long

  If KeyCode = KeyCodeConstants.vbKeyEscape Then
    TxtBoxU.GetDraggedTextRange firstChar, lastChar
    If (firstChar >= 0) And (lastChar >= 0) Then TxtBoxU.EndDrag True
  End If
End Sub

Private Sub Form_Load()
  Set ToolTip = New clsToolTip
  With ToolTip
    .Create Me.hWnd
    .Activate
    .AddTool chkOLEDragDrop.hWnd, "Check to start OLE Drag'n'Drop operations instead of normal ones.", , , False
  End With

  LoadTextU
End Sub

Private Sub Form_Terminate()
  Set ToolTip = Nothing
End Sub

Private Sub mnuCancel_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCancel
End Sub

Private Sub mnuCopy_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCopy
End Sub

Private Sub mnuMove_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciMove
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  suppressDefaultContextMenu = False
End Sub

Private Sub TxtBoxU_AbortedDrag()
  TxtBoxU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, -1
End Sub

Private Sub TxtBoxU_BeginDrag(ByVal firstChar As Long, ByVal lastChar As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  bRightDrag = False
  If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
    TxtBoxU.OLEDrag , , , firstChar, lastChar
  Else
    TxtBoxU.BeginDrag firstChar, lastChar, -1
  End If
End Sub

Private Sub TxtBoxU_BeginRDrag(ByVal firstChar As Long, ByVal lastChar As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  bRightDrag = True
  If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
    TxtBoxU.OLEDrag , , , firstChar, lastChar
  Else
    TxtBoxU.BeginDrag firstChar, lastChar, -1
  End If
End Sub

Private Sub TxtBoxU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  ' we don't want the default menu during right-button drag'n'drop
  showDefaultMenu = Not suppressDefaultContextMenu
End Sub

Private Sub TxtBoxU_DragMouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim newCharIndex As Long
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants

  TxtBoxU.GetClosestInsertMarkPosition ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), newInsertMarkRelativePosition, newCharIndex
  TxtBoxU.SetInsertMarkPosition newInsertMarkRelativePosition, newCharIndex
End Sub

Private Sub TxtBoxU_Drop(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim draggedText As String
  Dim firstDraggedChar As Long
  Dim insertAt As Long
  Dim insertionMark As InsertMarkPositionConstants
  Dim lastDraggedChar As Long
  Dim newText As String
  Dim oldText As String

  If bRightDrag Then
    x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    Me.PopupMenu mnuPopup, , x, y, mnuMove
    Select Case chosenMenuItem
      Case ChosenMenuItemConstants.ciCancel
        TxtBoxU_AbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciCopy
        ' copy the text
        TxtBoxU.GetInsertMarkPosition insertionMark, insertAt
        If insertionMark <> InsertMarkPositionConstants.impNowhere Then
          TxtBoxU.GetDraggedTextRange firstDraggedChar, lastDraggedChar
          oldText = TxtBoxU.Text
          draggedText = Mid$(oldText, firstDraggedChar + 1, lastDraggedChar - firstDraggedChar + 1)

          If insertionMark = InsertMarkPositionConstants.impAfter Then
            insertAt = insertAt + 1
          End If

          newText = Left$(oldText, insertAt)
          newText = newText & draggedText
          newText = newText & Mid$(oldText, insertAt + 1)
          TxtBoxU.Text = newText
          TxtBoxU.SetSelection insertAt, insertAt + (lastDraggedChar - firstDraggedChar + 1)
        End If

        TxtBoxU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, -1
        Exit Sub
      Case ChosenMenuItemConstants.ciMove
        ' fall through
    End Select
  End If

  ' move the text
  TxtBoxU.GetInsertMarkPosition insertionMark, insertAt
  If insertionMark <> InsertMarkPositionConstants.impNowhere Then
    TxtBoxU.GetDraggedTextRange firstDraggedChar, lastDraggedChar
    oldText = TxtBoxU.Text
    draggedText = Mid$(oldText, firstDraggedChar + 1, lastDraggedChar - firstDraggedChar + 1)

    If insertionMark = InsertMarkPositionConstants.impAfter Then
      insertAt = insertAt + 1
    End If

    If insertAt < firstDraggedChar Then
      ' ATTENTION: Left and Mid work one-based, but the control works zero-based!
      If insertAt > 0 Then
        newText = Left$(oldText, insertAt)
      End If
      newText = newText & draggedText
      newText = newText & Mid$(oldText, insertAt + 1, firstDraggedChar - insertAt)
      newText = newText & Mid$(oldText, lastDraggedChar + 2)
      TxtBoxU.Text = newText
      TxtBoxU.SetSelection insertAt, insertAt + (lastDraggedChar - firstDraggedChar + 1)
    ElseIf insertAt > lastDraggedChar + 1 Then
      ' ATTENTION: Left and Mid work one-based, but the control works zero-based!
      If firstDraggedChar > 0 Then
        newText = Left$(oldText, firstDraggedChar)
      End If
      newText = newText & Mid$(oldText, lastDraggedChar + 2, insertAt - lastDraggedChar - 1)
      newText = newText & draggedText
      newText = newText & Mid$(oldText, insertAt + 1)
      TxtBoxU.Text = newText
      TxtBoxU.SetSelection firstDraggedChar + (insertAt - lastDraggedChar - 1), insertAt
    Else
      ' restore the selection
      TxtBoxU.SetSelection firstDraggedChar, lastDraggedChar + 1
    End If
  End If

  TxtBoxU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, -1
End Sub

Private Sub TxtBoxU_KeyDown(KeyCode As Integer, ByVal shift As Integer)
  Dim firstChar As Long
  Dim lastChar As Long

  If KeyCode = KeyCodeConstants.vbKeyEscape Then
    TxtBoxU.GetDraggedTextRange firstChar, lastChar
    If (firstChar >= 0) And (lastChar >= 0) Then TxtBoxU.EndDrag True
  End If
End Sub

Private Sub TxtBoxU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim firstChar As Long
  Dim insertAt As Long
  Dim insertionMark As InsertMarkPositionConstants
  Dim lastChar As Long

  If chkOLEDragDrop.Value = CheckBoxConstants.vbUnchecked Then
    If ((button = MouseButtonConstants.vbLeftButton) And (Not bRightDrag)) Or ((button = MouseButtonConstants.vbRightButton) And bRightDrag) Then
      TxtBoxU.GetDraggedTextRange firstChar, lastChar
      If (firstChar >= 0) And (lastChar >= 0) Then
        TxtBoxU.GetInsertMarkPosition insertionMark, insertAt
        suppressDefaultContextMenu = True
        If insertionMark = InsertMarkPositionConstants.impNowhere Then
          ' cancel
          TxtBoxU.EndDrag True
        Else
          ' drop
          TxtBoxU.EndDrag False
        End If
        ' suppressing the default context menu is a bit tricky - use a timer to reset the flag
        Timer1.Enabled = True
      End If
    End If
  End If
End Sub

Private Sub TxtBoxU_OLECompleteDrag(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, ByVal performedEffect As EditCtlsLibUCtl.OLEDropEffectConstants)
  Dim bytArray() As Byte
  Dim CLSIDOfTarget As GUID
  Dim firstChar As Long
  Dim lastChar As Long

  If performedEffect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove Then
    ' remove the dragged text
    TxtBoxU.GetDraggedTextRange firstChar, lastChar
    TxtBoxU.SetSelection firstChar, lastChar + 1
    TxtBoxU.ReplaceSelectedText ""
  ElseIf Data.GetFormat(CFSTR_TARGETCLSID) Then
    bytArray = Data.GetData(CFSTR_TARGETCLSID)
    CopyMemory VarPtr(CLSIDOfTarget), VarPtr(bytArray(LBound(bytArray))), LenB(CLSIDOfTarget)
    If IsEqualGUID(CLSIDOfTarget, CLSID_RecycleBin) Then
      ' remove the dragged text
      TxtBoxU.GetDraggedTextRange firstChar, lastChar
      TxtBoxU.SetSelection firstChar, lastChar + 1
      TxtBoxU.ReplaceSelectedText ""
    End If
  End If
End Sub

Private Sub TxtBoxU_OLEDragDrop(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim bytArray() As Byte
  Dim files() As String
  Dim i As Long
  Dim insertAt As Long
  Dim insertionMark As InsertMarkPositionConstants
  Dim p As Long
  Dim str As String

  If Data.GetFormat(CF_UNICODETEXT) Then
    str = Data.GetData(CF_UNICODETEXT)
  ElseIf Data.GetFormat(vbCFText) Then
    str = Data.GetData(vbCFText)
  ElseIf Data.GetFormat(CF_OEMTEXT) Then
    str = Data.GetData(CF_OEMTEXT)
  ElseIf Data.GetFormat(vbCFFiles) Then
    ' insert a line for each file/folder
    On Error GoTo NoFiles
    files = Data.GetData(vbCFFiles)
    For i = LBound(files) To UBound(files)
      p = InStrRev(files(i), "\")
      If p > 0 Then files(i) = Mid$(files(i), p + 1)
      str = str & files(i) & vbNewLine
    Next i
NoFiles:
  End If

  If str <> "" Then
    ' insert the text

    TxtBoxU.GetInsertMarkPosition insertionMark, insertAt
    Select Case insertionMark
      Case InsertMarkPositionConstants.impAfter
        TxtBoxU.SetSelection insertAt + 1, insertAt + 1
        TxtBoxU.ReplaceSelectedText str
      Case InsertMarkPositionConstants.impBefore
        TxtBoxU.SetSelection insertAt, insertAt
        TxtBoxU.ReplaceSelectedText str
    End Select
  End If

  TxtBoxU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, -1

  ' Be careful with odeMove!! Some sources delete the data themselves if the Move effect was
  ' chosen. So you may lose data!
'  If shift And vbShiftMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeLink

'  If Data.GetFormat(vbCFText) Then MsgBox "Dropped Text: " & Data.GetData(vbCFText)
'  If Data.GetFormat(CF_OEMTEXT) Then MsgBox "Dropped OEM-Text: " & Data.GetData(7)
'  If Data.GetFormat(CF_UNICODETEXT) Then MsgBox "Dropped Unicode-Text: " & Data.GetData(13)
'  If Data.GetFormat(vbCFRTF) Then MsgBox "Dropped RTF-Text: " & Data.GetData(vbCFRTF)
'  If Data.GetFormat(CFSTR_HTML) Then MsgBox "Dropped HTML: " & Data.GetData(CFSTR_HTML)
'  If Data.GetFormat(CFSTR_HTML2) Then MsgBox "Dropped HTML2: " & Data.GetData(CFSTR_HTML2)
'  If Data.GetFormat(CFSTR_MOZILLAHTMLCONTEXT) Then MsgBox "Dropped text/_moz_htmlcontext: " & Data.GetData(CFSTR_MOZILLAHTMLCONTEXT)
'  If Data.GetFormat(CFSTR_TEXTHTML) Then MsgBox "Dropped text/html: " & Data.GetData(CFSTR_TEXTHTML)
'  If data.GetFormat(vbCFFiles) Then
'    str = ""
'    For Each file In data.Files
'      str = str & file & vbNewLine
'    Next file
'    MsgBox "Dropped files:" & vbNewLine & str
'  End If
'
'  ReDim bytArray(1 To LenB(CLSID_RecycleBin)) As Byte
'  CopyMemory VarPtr(bytArray(LBound(bytArray))), VarPtr(CLSID_RecycleBin), LenB(CLSID_RecycleBin)
'  Data.SetData CFSTR_TARGETCLSID, bytArray
'
'  ' another way to inform the source of the performed drop effect
'  Data.SetData CFSTR_PERFORMEDDROPEFFECT, Effect
End Sub

Private Sub TxtBoxU_OLEDragEnter(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim newCharIndex As Long
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants

  TxtBoxU.GetClosestInsertMarkPosition ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), newInsertMarkRelativePosition, newCharIndex
  TxtBoxU.SetInsertMarkPosition newInsertMarkRelativePosition, newCharIndex

  effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbShiftMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeLink
End Sub

Private Sub TxtBoxU_OLEDragLeave(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  TxtBoxU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, -1
End Sub

Private Sub TxtBoxU_OLEDragMouseMove(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim newCharIndex As Long
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants

  TxtBoxU.GetClosestInsertMarkPosition ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), newInsertMarkRelativePosition, newCharIndex
  TxtBoxU.SetInsertMarkPosition newInsertMarkRelativePosition, newCharIndex

  effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbShiftMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeLink
End Sub

Private Sub TxtBoxU_OLESetData(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal index As Long, ByVal dataOrViewAspect As Long)
  Dim firstChar As Long
  Dim lastChar As Long

  Select Case formatID
    Case vbCFText, CF_OEMTEXT, CF_UNICODETEXT
      TxtBoxU.GetDraggedTextRange firstChar, lastChar
      Data.SetData formatID, Mid$(TxtBoxU.Text, firstChar + 1, lastChar - firstChar + 1)
  End Select
End Sub

Private Sub TxtBoxU_OLEStartDrag(ByVal Data As EditCtlsLibUCtl.IOLEDataObject)
  Data.SetData vbCFText
  Data.SetData CF_OEMTEXT
  Data.SetData CF_UNICODETEXT
End Sub

Private Sub TxtBoxU_RecreatedControlWindow(ByVal hWnd As Long)
  LoadTextU
End Sub


Private Sub LoadTextU()
  Dim filePath As String
  Dim fn As Long

  If Right$(App.path, 3) = "bin" Then
    filePath = App.path & "\clsToolTip.cls"
  Else
    filePath = App.path & "\bin\clsToolTip.cls"
  End If
  fn = FreeFile
  On Error GoTo FileNotFound
  Open filePath For Input As #fn
    TxtBoxU.Text = Input(LOF(fn), fn)
  Close #fn
  Exit Sub

FileNotFound:
End Sub

Private Function IsEqualGUID(IID1 As GUID, IID2 As GUID) As Boolean
  Dim Tmp1 As Currency
  Dim Tmp2 As Currency

  If IID1.Data1 = IID2.Data1 Then
    If IID1.Data2 = IID2.Data2 Then
      If IID1.Data3 = IID2.Data3 Then
        CopyMemory VarPtr(Tmp1), VarPtr(IID1.Data4(0)), 8
        CopyMemory VarPtr(Tmp2), VarPtr(IID2.Data4(0)), 8
        IsEqualGUID = (Tmp1 = Tmp2)
      End If
    End If
  End If
End Function
