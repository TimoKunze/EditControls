// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:A8F9B8E7-E699-4fce-A647-72C877F8E632> version("1.10") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_EDITCTLSEDIT, CMainDlg, &__uuidof(_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> editctlsTextBoxContainerWnd;
	CContainedWindowT<CAxWindow> editctlsTextBoxWnd;

	CMainDlg() :
	    editctlsTextBoxContainerWnd(this, 1),
	    editctlsTextBoxWnd(this, 2)
	{
	}

	struct Controls
	{
		CStatic descrStatic;
		CEdit countEdit;
		CUpDownCtrl countSpin;
		CButton fillTextBoxesButton;
		CStatic fillTimeStatic;
		CButton aboutButton;
		CToolTipCtrl tooltip;
		CComPtr<ITextBox> editctlsTextBox;
		CEdit nativeEdit;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_HANDLER(IDC_FILLTEXTBOXES, BN_CLICKED, OnBnClickedFillTextBoxesButton)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnBnClickedFillTextBoxesButton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
};
