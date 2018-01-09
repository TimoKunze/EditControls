// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include <olectl.h>
#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.editctlsTextBox) {
		DispEventUnadvise(controls.editctlsTextBox);
		controls.editctlsTextBox.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.descrStatic = GetDlgItem(IDC_DESCRSTATIC);
	controls.countEdit = GetDlgItem(IDC_LINESEDIT);
	controls.countSpin = GetDlgItem(IDC_LINESSPIN);
	controls.countSpin.SetRange32(1, 2000);
	controls.countSpin.SetPos32(1000);
	controls.fillTextBoxesButton = GetDlgItem(IDC_FILLTEXTBOXES);
	controls.fillTimeStatic = GetDlgItem(IDC_FILLTIMESTATIC);
	controls.fillTimeStatic.SetWindowText(TEXT("Fill time: 0 ms/0 ms"));
	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);
	controls.tooltip.Create(*this, 0, NULL, WS_OVERLAPPED | WS_POPUP | WS_VISIBLE | TTS_NOPREFIX | TTS_BALLOON, WS_EX_LEFT | WS_EX_LTRREADING | WS_EX_TOOLWINDOW);
	controls.tooltip.Activate(TRUE);

	editctlsTextBoxContainerWnd.SubclassWindow(GetDlgItem(IDC_EDITCTLSEDIT));
	editctlsTextBoxContainerWnd.QueryControl(__uuidof(ITextBox), reinterpret_cast<LPVOID*>(&controls.editctlsTextBox));
	if(controls.editctlsTextBox) {
		DispEventAdvise(controls.editctlsTextBox);
		editctlsTextBoxWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.editctlsTextBox->GethWnd())));
		CToolInfo ti(TTF_SUBCLASS | TTF_CENTERTIP, editctlsTextBoxWnd, 0, NULL, TEXT("EditControls TextBox"));
		controls.tooltip.AddTool(ti);
	}
	controls.nativeEdit = GetDlgItem(IDC_NATIVEEDIT);
	if(controls.nativeEdit) {
		CToolInfo ti(TTF_SUBCLASS | TTF_CENTERTIP, controls.nativeEdit, 0, NULL, TEXT("Native TextBox"));
		controls.tooltip.AddTool(ti);
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.aboutButton.IsWindow()) {
		WINDOWPOS* pDetails = (WINDOWPOS*) lParam;

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			WTL::CRect clientRectangle;
			GetClientRect(&clientRectangle);
			controls.aboutButton.SetWindowPos(NULL, clientRectangle.Width() - 80, clientRectangle.Height() - 28, 0, 0, SWP_NOSIZE);
			controls.descrStatic.SetWindowPos(NULL, 5, clientRectangle.Height() - 23, 0, 0, SWP_NOSIZE);
			controls.countEdit.SetWindowPos(NULL, 38, clientRectangle.Height() - 25, 65, 17, 0);
			controls.countSpin.SetBuddy(controls.countEdit);
			controls.fillTextBoxesButton.SetWindowPos(NULL, 108, clientRectangle.Height() - 28, 0, 0, SWP_NOSIZE);
			controls.fillTimeStatic.SetWindowPos(NULL, 210, clientRectangle.Height() - 23, 0, 0, SWP_NOSIZE);

			editctlsTextBoxContainerWnd.SetWindowPos(NULL, 0, 0, clientRectangle.Width() / 2 - 3, clientRectangle.Height() - 66, SWP_NOMOVE);
			controls.nativeEdit.SetWindowPos(NULL, clientRectangle.Width() / 2 + 3, 33, clientRectangle.Width() / 2 - 3, clientRectangle.Height() - 66, 0);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnBnClickedFillTextBoxesButton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.fillTextBoxesButton.EnableWindow(FALSE);
	controls.editctlsTextBox->PutText(_bstr_t(TEXT("")).Detach());
	controls.nativeEdit.SetWindowText(TEXT(""));
	int cLines = controls.countSpin.GetPos32();
	CAtlString str;

	DWORD dwStartEditCtls = GetTickCount();
	for(int iLine = 1; iLine <= cLines; ++iLine) {
		str.Format(TEXT("This is line %i\r\n"), iLine);
		controls.editctlsTextBox->AppendText(_bstr_t(str).Detach(), VARIANT_TRUE, VARIANT_TRUE);
	}
	DWORD dwEndEditCtls = GetTickCount();

	DWORD dwStartNative = GetTickCount();
	for(int iLine = 1; iLine <= cLines; ++iLine) {
		str.Format(TEXT("This is line %i\r\n"), iLine);
		controls.nativeEdit.AppendText(str);
		controls.nativeEdit.SetSel(INT_MAX, INT_MAX);
	}
	DWORD dwEndNative = GetTickCount();

	str.Format(TEXT("Fill time: %i ms/%i ms"), dwEndEditCtls - dwStartEditCtls, dwEndNative - dwStartNative);
	controls.fillTimeStatic.SetWindowText(str);
	controls.fillTextBoxesButton.EnableWindow(TRUE);
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.editctlsTextBox) {
		controls.editctlsTextBox->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}