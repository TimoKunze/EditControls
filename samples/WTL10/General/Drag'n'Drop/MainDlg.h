// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:A8F9B8E7-E699-4fce-A647-72C877F8E632> version("1.11") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_TXTBOXU, CMainDlg, &__uuidof(_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 11>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> txtBoxUWnd;

	CMainDlg() :
	    txtBoxUWnd(this, 1)
	{
		CF_TARGETCLSID = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_TARGETCLSID));
		suppressDefaultContextMenu = FALSE;
	}

	BOOL bRightDrag;
	BOOL suppressDefaultContextMenu;
	#define TIMERID_CONTEXTMENUHACK 1

	struct Controls
	{
		CButton oleDDCheck;
		CComPtr<ITextBox> txtboxU;
	} controls;

	CLIPFORMAT CF_TARGETCLSID;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 1, AbortedDragTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 3, BeginDragTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 4, BeginRDragTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 6, ContextMenuTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 9, DragMouseMoveTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 10, DropTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 11, KeyDownTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 21, MouseUpTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 22, OLECompleteDragTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 23, OLEDragDropTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 24, OLEDragEnterTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 25, OLEDragLeaveTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 26, OLEDragMouseMoveTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 29, OLESetDataTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 30, OLEStartDragTxtboxu)
		SINK_ENTRY_EX(IDC_TXTBOXU, __uuidof(_ITextBoxEvents), 34, RecreatedControlWindowTxtboxu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_TXTBOXU, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_OLEDDCHECK, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void LoadText(void);

	void __stdcall AbortedDragTxtboxu();
	void __stdcall BeginDragTxtboxu(LONG firstChar, LONG lastChar, short /*button*/, short /*shift*/, long /*x*/, long /*y*/);
	void __stdcall BeginRDragTxtboxu(LONG firstChar, LONG lastChar, short /*button*/, short /*shift*/, long /*x*/, long /*y*/);
	void __stdcall ContextMenuTxtboxu(short /*button*/, short /*shift*/, long /*x*/, long /*y*/, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DragMouseMoveTxtboxu(short /*button*/, short /*shift*/, long x, long y, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall DropTxtboxu(short /*button*/, short /*shift*/, long x, long y);
	void __stdcall KeyDownTxtboxu(short* keyCode, short /*shift*/);
	void __stdcall MouseUpTxtboxu(short button, short /*shift*/, long /*x*/, long /*y*/);
	void __stdcall OLECompleteDragTxtboxu(LPDISPATCH data, long performedEffect);
	void __stdcall OLEDragDropTxtboxu(LPDISPATCH data, long* effect, short /*button*/, short shift, long /*x*/, long /*y*/);
	void __stdcall OLEDragEnterTxtboxu(LPDISPATCH /*data*/, long* effect, short /*button*/, short shift, long x, long y, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall OLEDragLeaveTxtboxu(LPDISPATCH /*data*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/);
	void __stdcall OLEDragMouseMoveTxtboxu(LPDISPATCH /*data*/, long* effect, short /*button*/, short shift, long x, long y, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall OLESetDataTxtboxu(LPDISPATCH data, long formatID, long /*index*/, long /*dataOrViewAspect*/);
	void __stdcall OLEStartDragTxtboxu(LPDISPATCH data);
	void __stdcall RecreatedControlWindowTxtboxu(long /*hWnd*/);
};
