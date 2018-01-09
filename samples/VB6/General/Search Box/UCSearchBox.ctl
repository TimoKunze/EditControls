VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.10#0"; "EditCtlsU.ocx"
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.0#0"; "TBarCtlsU.ocx"
Begin VB.UserControl UCSearchBox 
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   Begin TBarCtlsLibUCtl.ToolBar SearchButton 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   495
      _cx             =   873
      _cy             =   450
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   0   'False
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   0
      BorderStyle     =   0
      ButtonHeight    =   28
      ButtonStyle     =   1
      ButtonTextPosition=   1
      ButtonWidth     =   22
      DisabledEvents  =   917994
      DisplayMenuDivider=   0   'False
      DisplayPartiallyClippedButtons=   0   'False
      DontRedraw      =   0   'False
      DragClickTime   =   -1
      DragDropCustomizationModifierKey=   0
      DropDownGap     =   -1
      Enabled         =   -1  'True
      FirstButtonIndentation=   0
      FocusOnClick    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   -1
      HorizontalButtonPadding=   -1
      HorizontalButtonSpacing=   0
      HorizontalIconCaptionGap=   -1
      HorizontalTextAlignment=   0
      HoverTime       =   -1
      InsertMarkColor =   0
      MaximumButtonWidth=   0
      MaximumTextRows =   1
      MenuBarTheme    =   1
      MenuMode        =   0   'False
      MinimumButtonWidth=   0
      MousePointer    =   0
      MultiColumn     =   0   'False
      NormalDropDownButtonStyle=   0
      OLEDragImageStyle=   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   -1  'True
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      ShowToolTips    =   -1  'True
      SupportOLEDragImages=   -1  'True
      UseMnemonics    =   -1  'True
      UseSystemFont   =   -1  'True
      VerticalButtonPadding=   -1
      VerticalButtonSpacing=   0
      VerticalTextAlignment=   1
      WrapButtons     =   0   'False
   End
   Begin EditCtlsLibUCtl.TextBox txtSearch 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2415
      _cx             =   4260
      _cy             =   556
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   0
      AutoScrolling   =   2
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
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
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
      CueBanner       =   "UCSearchBox.ctx":0000
      Text            =   "UCSearchBox.ctx":0020
   End
End
Attribute VB_Name = "UCSearchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  ' This sample was heavily inspired by the "Explorer breadcrumbs in Vista" sample by
  ' Bjarke Viksoe which can be found on http://www.viksoe.dk/code/breadcrumbs.htm


  Implements ISubclassedWindow


  Private Const BTNID_SEARCH As Long = 1


  Private m_hotTracked As Long
  Private m_hImageList As Long
  Private m_hTheme As Long
  Private m_rcButton As RECT
  Private m_subclassed As Boolean


  Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hrgnDest As Long, ByVal hrgnSrc1 As Long, ByVal hrgnSrc2 As Long, ByVal fnCombineMode As Long) As Long
  Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (lprc As RECT) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal pClipRect As Long) As Long
  Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal hDC As Long, prc As RECT) As Long
  Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function GetFocusAPI Lib "user32.dll" Alias "GetFocus" () As Long
  Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetThemeInt Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByRef piVal As Long) As Long
  Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Private Declare Function ImageList_LoadImage Lib "comctl32.dll" Alias "ImageList_LoadImageW" (ByVal hi As Long, ByVal lpbmp As Long, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function OffsetRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long


  Public Event ApplyFilter()
  Public Event ClearFilter()
  Public Event FilterChange()


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidUCSearchBox
      lRet = HandleMessage_UserControl(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "UCSearchBox.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in UCSearchBox.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_UserControl(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const COLOR_WINDOW As Long = 5
  Const RGN_AND As Long = 1
  Const SM_CXBORDER As Long = 5
  Const SM_CXEDGE As Long = 45
  Const SM_CYBORDER As Long = 6
  Const SM_CYEDGE As Long = 46
  Const SWP_NOACTIVATE As Long = &H10
  Const SWP_NOZORDER As Long = &H4
  Const TMT_SIZINGBORDERWIDTH As Long = 1201
  Const WM_ERASEBKGND As Long = &H14
  Const WM_NCPAINT As Long = &H85
  Const WM_PAINT As Long = &HF
  Const WM_PRINTCLIENT As Long = &H318
  Const WM_SIZE As Long = &H5
  Const WM_THEMECHANGED As Long = &H31A
  Dim cxBorder As Long
  Dim cxEdge As Long
  Dim cyBorder As Long
  Dim cyEdge As Long
  Dim hDC As Long
  Dim hRgn As Long
  Dim hUpdateRgn As Long
  Dim ps As PAINTSTRUCT
  Dim rcClient As RECT
  Dim rcEdit As RECT
  Dim rcToolBar As RECT
  Dim rcWnd As RECT
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_ERASEBKGND
      ' make sure the control's background is drawn properly
      If Not IsCompositionEnabled Then DoPaint wParam
      lRet = 1
      bCallDefProc = False
    Case WM_NCPAINT
      ' draws the control's border
      hUpdateRgn = IIf(wParam <> 1, wParam, 0)
      hDC = GetWindowDC(hWnd)
      If hDC Then
        cxBorder = GetSystemMetrics(SM_CXBORDER)
        cyBorder = GetSystemMetrics(SM_CYBORDER)
        If GetThemeInt(m_hTheme, 0, 0, TMT_SIZINGBORDERWIDTH, cxBorder) >= S_OK Then
          cyBorder = cxBorder
        End If

        GetWindowRect hWnd, rcWnd
        cxEdge = GetSystemMetrics(SM_CXEDGE)
        cyEdge = GetSystemMetrics(SM_CYEDGE)
        InflateRect rcWnd, -cxEdge, -cyEdge
        hRgn = CreateRectRgnIndirect(rcWnd)
        If hRgn Then
          If hUpdateRgn Then
            CombineRgn hRgn, hUpdateRgn, hRgn, RGN_AND
          End If
          OffsetRect rcWnd, -rcWnd.Left, -rcWnd.Top
          OffsetRect rcWnd, cxEdge, cyEdge
          ExcludeClipRect hDC, rcWnd.Left, rcWnd.Top, rcWnd.Right, rcWnd.Bottom
          InflateRect rcWnd, cxEdge, cyEdge

          If (cxBorder < cxEdge) And (cyBorder < cyEdge) Then
            InflateRect rcWnd, cxBorder - cxEdge, cyBorder - cyEdge
            FillRect hDC, rcWnd, GetSysColorBrush(COLOR_WINDOW)
          End If
          DefSubclassProc hWnd, WM_NCPAINT, hRgn, 0

          DeleteObject hRgn
        End If

        ReleaseDC hWnd, hDC
      End If
      bCallDefProc = False
    Case WM_PAINT, WM_PRINTCLIENT
      ' make sure the control's background is drawn properly
      If wParam Then
        DoPaint wParam
      Else
        hDC = BeginPaint(hWnd, ps)
        If IsCompositionEnabled Then DoPaint hDC
        EndPaint hWnd, ps
      End If
      bCallDefProc = False
    Case WM_SIZE
      ' reposition the text box and the tool bar
      rcClient.Right = LoWord(lParam)
      rcClient.Bottom = HiWord(lParam)
      rcEdit.Top = 4
      rcEdit.Right = (rcClient.Right - rcClient.Left) - (m_rcButton.Right - m_rcButton.Left)
      rcEdit.Bottom = rcClient.Bottom - 4
      rcToolBar.Left = (rcClient.Right - rcClient.Left) - (m_rcButton.Right - m_rcButton.Left) - 4
      rcToolBar.Top = 1
      rcToolBar.Right = rcClient.Right - 2
      rcToolBar.Bottom = rcClient.Bottom - 1
      SetWindowPos txtSearch.hWnd, 0, rcEdit.Left, rcEdit.Top, rcEdit.Right - rcEdit.Left, rcEdit.Bottom - rcEdit.Top, SWP_NOACTIVATE Or SWP_NOZORDER
      SetWindowPos SearchButton.hWnd, 0, rcToolBar.Left, rcToolBar.Top, rcToolBar.Right - rcToolBar.Left, rcToolBar.Bottom - rcToolBar.Top, SWP_NOACTIVATE Or SWP_NOZORDER
      bCallDefProc = False
    Case WM_THEMECHANGED
      CloseThemeData m_hTheme
      m_hTheme = OpenThemeData(hWnd, StrPtr("SearchBoxComposited::SearchBox"))
  End Select

StdHandler_Ende:
  HandleMessage_UserControl = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in UCSearchBox.HandleMessage_UserControl: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub SearchButton_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  If commandID = BTNID_SEARCH Then
    If txtSearch.TextLength > 0 Then
      txtSearch.Text = ""
      RaiseEvent ClearFilter
    Else
      RaiseEvent ApplyFilter
    End If
  End If
End Sub

Private Sub SearchButton_MouseEnter(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  m_hotTracked = m_hotTracked + 1
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub SearchButton_MouseLeave(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  m_hotTracked = m_hotTracked - 1
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub txtSearch_GotFocus()
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub txtSearch_KeyDown(keyCode As Integer, ByVal shift As Integer)
  If keyCode = vbKeyReturn Then
    RaiseEvent ApplyFilter
  End If
End Sub

Private Sub txtSearch_LostFocus()
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub txtSearch_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  m_hotTracked = m_hotTracked + 1
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub txtSearch_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  m_hotTracked = m_hotTracked - 1
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub txtSearch_TextChanged()
  ' change icon
  SearchButton.Buttons(BTNID_SEARCH, btitID).IconIndex = IIf(txtSearch.TextLength > 0, 1, 0)
  RaiseEvent FilterChange
End Sub

Private Sub UserControl_GotFocus()
  txtSearch.SetFocus
  UserControl.Refresh
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub UserControl_Initialize()
  Const IMAGE_BITMAP As Long = 0
  Const LR_CREATEDIBSECTION As Long = &H2000
  Const LR_LOADFROMFILE As Long = &H10
  Dim btn As ToolBarButton
  Dim imgDir As String

  m_hTheme = OpenThemeData(UserControl.hWnd, StrPtr("SearchBoxComposited::SearchBox"))

  If Right$(App.Path, 3) = "bin" Then
    imgDir = App.Path & "\..\res\"
  Else
    imgDir = App.Path & "\res\"
  End If
  m_hImageList = ImageList_LoadImage(0, StrPtr(imgDir & "SearchButton.bmp"), IMAGE_BITMAP, 16, 0, vbBlack, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
  SearchButton.hImageList(ilNormalButtons) = m_hImageList

  Set btn = SearchButton.Buttons.Add(BTNID_SEARCH, , 0, AutoSize:=False)
  If Not (btn Is Nothing) Then
    btn.GetRectangle brtButton, m_rcButton.Left, m_rcButton.Top, m_rcButton.Right, m_rcButton.Bottom
  End If

  ' set some special themes for the text box and the tool bar
  SetWindowTheme txtSearch.hWnd, StrPtr("SearchBoxEditComposited"), 0
  SetWindowTheme SearchButton.hWnd, StrPtr("SearchButtonComposited"), 0
End Sub

Private Sub UserControl_InitProperties()
  txtSearch.CueBanner = "Search"
End Sub

Private Sub UserControl_LostFocus()
  UserControl.Refresh
  txtSearch.Refresh
  SearchButton.Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    CueBanner = PropBag.ReadProperty("CueBanner", "Search")
    Enabled = PropBag.ReadProperty("Enabled", True)
    Filter = PropBag.ReadProperty("Filter", "")
  End With
End Sub

Private Sub UserControl_Resize()
  Const SWP_NOACTIVATE As Long = &H10
  Const SWP_NOZORDER As Long = &H4
  Dim rcClient As RECT
  Dim rcEdit As RECT
  Dim rcToolBar As RECT

  If IsInIDE Then
    ' during runtime the child controls will be aligned whenever we receive a WM_SIZE message
    GetClientRect UserControl.hWnd, rcClient
    rcEdit.Top = 4
    rcEdit.Right = (rcClient.Right - rcClient.Left) - (m_rcButton.Right - m_rcButton.Left)
    rcEdit.Bottom = rcClient.Bottom - 4
    rcToolBar.Left = (rcClient.Right - rcClient.Left) - (m_rcButton.Right - m_rcButton.Left) - 4
    rcToolBar.Top = 1
    rcToolBar.Right = rcClient.Right - 2
    rcToolBar.Bottom = rcClient.Bottom - 1
    SetWindowPos txtSearch.hWnd, 0, rcEdit.Left, rcEdit.Top, rcEdit.Right - rcEdit.Left, rcEdit.Bottom - rcEdit.Top, SWP_NOACTIVATE Or SWP_NOZORDER
    SetWindowPos SearchButton.hWnd, 0, rcToolBar.Left, rcToolBar.Top, rcToolBar.Right - rcToolBar.Left, rcToolBar.Bottom - rcToolBar.Top, SWP_NOACTIVATE Or SWP_NOZORDER
  ElseIf Not m_subclassed Then
    If SubclassWindow(UserControl.hWnd, Me, EnumSubclassID.escidUCSearchBox) Then
      m_subclassed = True
    Else
      Debug.Print "Subclassing failed!"
    End If
  End If
End Sub

Private Sub UserControl_Terminate()
  If UnSubclassWindow(UserControl.hWnd, EnumSubclassID.escidUCSearchBox) Then
    m_subclassed = False
  Else
    Debug.Print "UnSubclassing failed!"
  End If
  If m_hTheme Then
    CloseThemeData m_hTheme
  End If
  If m_hImageList Then
    ImageList_Destroy m_hImageList
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    PropBag.WriteProperty "CueBanner", CueBanner, "Search"
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Filter", Filter, ""
  End With
End Sub


Public Property Get CueBanner() As String
  CueBanner = txtSearch.CueBanner
End Property

Public Property Let CueBanner(ByVal newValue As String)
  If txtSearch.CueBanner <> newValue Then
    txtSearch.CueBanner = newValue
    PropertyChanged "CueBanner"
  End If
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
  If UserControl.Enabled <> newValue Then
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
  End If
End Property

Public Property Get Filter() As String
  Filter = txtSearch.Text
End Property

Public Property Let Filter(ByVal newValue As String)
  If txtSearch.Text <> newValue Then
    txtSearch.Text = newValue
    PropertyChanged "Filter"
  End If
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property


Private Sub DoPaint(ByVal hDC As Long)
  Const SBB_DISABLED As Long = 3
  Const SBB_FOCUSED As Long = 4
  Const SBB_HOT As Long = 2
  Const SBB_NORMAL As Long = 1
  Const SBP_SBBACKGROUND As Long = 1
  Dim rcClient As RECT
  Dim stateID As Long

  ' draw the control background using the theme API
  GetClientRect UserControl.hWnd, rcClient
  stateID = SBB_NORMAL
  If m_hotTracked > 0 Then stateID = SBB_HOT
  If GetFocusAPI = txtSearch.hWnd Then stateID = SBB_FOCUSED
  If Not Me.Enabled Then stateID = SBB_DISABLED
  DrawThemeParentBackground UserControl.hWnd, hDC, rcClient
  DrawThemeBackground m_hTheme, hDC, SBP_SBBACKGROUND, stateID, rcClient, 0
End Sub

Private Function IsInIDE() As Boolean
  IsInIDE = Not Ambient.UserMode
End Function
