// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"


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
	if(controls.txtbox0U) {
		IDispEventImpl<IDC_TXTBOX0U, CMainDlg, &__uuidof(EditCtlsLibU::_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventUnadvise(controls.txtbox0U);
		controls.txtbox0U.Release();
	}
	if(controls.txtbox1U) {
		IDispEventImpl<IDC_TXTBOX1U, CMainDlg, &__uuidof(EditCtlsLibU::_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventUnadvise(controls.txtbox1U);
		controls.txtbox1U.Release();
	}
	if(controls.ipaddrU) {
		IDispEventImpl<IDC_IPADDRU, CMainDlg, &__uuidof(EditCtlsLibU::_IIPAddressBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventUnadvise(controls.ipaddrU);
		controls.ipaddrU.Release();
	}
	if(controls.hkboxU) {
		IDispEventImpl<IDC_HKBOXU, CMainDlg, &__uuidof(EditCtlsLibU::_IHotKeyBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventUnadvise(controls.hkboxU);
		controls.hkboxU.Release();
	}
	if(controls.udtxtboxU) {
		IDispEventImpl<IDC_UDTXTBOXU, CMainDlg, &__uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventUnadvise(controls.udtxtboxU);
		controls.udtxtboxU.Release();
	}
	if(controls.txtbox0A) {
		IDispEventImpl<IDC_TXTBOX0A, CMainDlg, &__uuidof(EditCtlsLibA::_ITextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventUnadvise(controls.txtbox0A);
		controls.txtbox0A.Release();
	}
	if(controls.txtbox1A) {
		IDispEventImpl<IDC_TXTBOX1A, CMainDlg, &__uuidof(EditCtlsLibA::_ITextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventUnadvise(controls.txtbox1A);
		controls.txtbox1A.Release();
	}
	if(controls.ipaddrA) {
		IDispEventImpl<IDC_IPADDRA, CMainDlg, &__uuidof(EditCtlsLibA::_IIPAddressBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventUnadvise(controls.ipaddrA);
		controls.ipaddrA.Release();
	}
	if(controls.hkboxA) {
		IDispEventImpl<IDC_HKBOXA, CMainDlg, &__uuidof(EditCtlsLibA::_IHotKeyBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventUnadvise(controls.hkboxA);
		controls.hkboxA.Release();
	}
	if(controls.udtxtboxA) {
		IDispEventImpl<IDC_UDTXTBOXA, CMainDlg, &__uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventUnadvise(controls.udtxtboxA);
		controls.udtxtboxA.Release();
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

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	aboutButton.SubclassWindow(GetDlgItem(ID_APP_ABOUT));

	txtbox0UWnd.SubclassWindow(GetDlgItem(IDC_TXTBOX0U));
	txtbox0UWnd.QueryControl(__uuidof(EditCtlsLibU::ITextBox), reinterpret_cast<LPVOID*>(&controls.txtbox0U));
	if(controls.txtbox0U) {
		IDispEventImpl<IDC_TXTBOX0U, CMainDlg, &__uuidof(EditCtlsLibU::_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventAdvise(controls.txtbox0U);
	}
	txtbox1UWnd.SubclassWindow(GetDlgItem(IDC_TXTBOX1U));
	txtbox1UWnd.QueryControl(__uuidof(EditCtlsLibU::ITextBox), reinterpret_cast<LPVOID*>(&controls.txtbox1U));
	if(controls.txtbox1U) {
		IDispEventImpl<IDC_TXTBOX1U, CMainDlg, &__uuidof(EditCtlsLibU::_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventAdvise(controls.txtbox1U);
	}
	ipaddrUWnd.SubclassWindow(GetDlgItem(IDC_IPADDRU));
	ipaddrUWnd.QueryControl(__uuidof(EditCtlsLibU::IIPAddressBox), reinterpret_cast<LPVOID*>(&controls.ipaddrU));
	if(controls.ipaddrU) {
		IDispEventImpl<IDC_IPADDRU, CMainDlg, &__uuidof(EditCtlsLibU::_IIPAddressBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventAdvise(controls.ipaddrU);
	}
	hkboxUWnd.SubclassWindow(GetDlgItem(IDC_HKBOXU));
	hkboxUWnd.QueryControl(__uuidof(EditCtlsLibU::IHotKeyBox), reinterpret_cast<LPVOID*>(&controls.hkboxU));
	if(controls.hkboxU) {
		IDispEventImpl<IDC_HKBOXU, CMainDlg, &__uuidof(EditCtlsLibU::_IHotKeyBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventAdvise(controls.hkboxU);
	}
	udtxtboxUWnd.SubclassWindow(GetDlgItem(IDC_UDTXTBOXU));
	udtxtboxUWnd.QueryControl(__uuidof(EditCtlsLibU::IUpDownTextBox), reinterpret_cast<LPVOID*>(&controls.udtxtboxU));
	if(controls.udtxtboxU) {
		IDispEventImpl<IDC_UDTXTBOXU, CMainDlg, &__uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>::DispEventAdvise(controls.udtxtboxU);
	}

	txtbox0AWnd.SubclassWindow(GetDlgItem(IDC_TXTBOX0A));
	txtbox0AWnd.QueryControl(__uuidof(EditCtlsLibA::ITextBox), reinterpret_cast<LPVOID*>(&controls.txtbox0A));
	if(controls.txtbox0A) {
		IDispEventImpl<IDC_TXTBOX0A, CMainDlg, &__uuidof(EditCtlsLibA::_ITextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventAdvise(controls.txtbox0A);
	}
	txtbox1AWnd.SubclassWindow(GetDlgItem(IDC_TXTBOX1A));
	txtbox1AWnd.QueryControl(__uuidof(EditCtlsLibA::ITextBox), reinterpret_cast<LPVOID*>(&controls.txtbox1A));
	if(controls.txtbox1A) {
		IDispEventImpl<IDC_TXTBOX1A, CMainDlg, &__uuidof(EditCtlsLibA::_ITextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventAdvise(controls.txtbox1A);
	}
	ipaddrAWnd.SubclassWindow(GetDlgItem(IDC_IPADDRA));
	ipaddrAWnd.QueryControl(__uuidof(EditCtlsLibA::IIPAddressBox), reinterpret_cast<LPVOID*>(&controls.ipaddrA));
	if(controls.ipaddrA) {
		IDispEventImpl<IDC_IPADDRA, CMainDlg, &__uuidof(EditCtlsLibA::_IIPAddressBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventAdvise(controls.ipaddrA);
	}
	hkboxAWnd.SubclassWindow(GetDlgItem(IDC_HKBOXA));
	hkboxAWnd.QueryControl(__uuidof(EditCtlsLibA::IHotKeyBox), reinterpret_cast<LPVOID*>(&controls.hkboxA));
	if(controls.hkboxA) {
		IDispEventImpl<IDC_HKBOXA, CMainDlg, &__uuidof(EditCtlsLibA::_IHotKeyBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventAdvise(controls.hkboxA);
	}
	udtxtboxAWnd.SubclassWindow(GetDlgItem(IDC_UDTXTBOXA));
	udtxtboxAWnd.QueryControl(__uuidof(EditCtlsLibA::IUpDownTextBox), reinterpret_cast<LPVOID*>(&controls.udtxtboxA));
	if(controls.udtxtboxA) {
		IDispEventImpl<IDC_UDTXTBOXA, CMainDlg, &__uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>::DispEventAdvise(controls.udtxtboxA);
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.logEdit.IsWindow() && aboutButton.IsWindow()) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			int cx = static_cast<int>(0.4 * static_cast<double>(clientRectangle.Width()));
			controls.logEdit.SetWindowPos(NULL, clientRectangle.Width() - cx, 0, cx, clientRectangle.Height() - 32, 0);
			aboutButton.SetWindowPos(NULL, clientRectangle.Width() - cx, clientRectangle.Height() - 27, 0, 0, SWP_NOSIZE);

			int halfHeight = (clientRectangle.Height() - 5) / 2;
			CRect remainingRectangle;
			controls.logEdit.GetWindowRect(&remainingRectangle);
			ScreenToClient(&remainingRectangle);
			remainingRectangle.right = remainingRectangle.left - 5;
			remainingRectangle.left = clientRectangle.left;
			remainingRectangle.top = clientRectangle.top;
			remainingRectangle.bottom = clientRectangle.bottom;

			txtbox0UWnd.SetWindowPos(NULL, 0, 0, remainingRectangle.Width(), halfHeight - 49, SWP_NOMOVE);
			txtbox1UWnd.SetWindowPos(NULL, 0, halfHeight - 44, remainingRectangle.Width(), 18, 0);
			ipaddrUWnd.SetWindowPos(NULL, 0, halfHeight - 21, 0, 0, SWP_NOSIZE);
			RECT rc = {0};
			hkboxUWnd.GetWindowRect(&rc);
			ScreenToClient(&rc);
			hkboxUWnd.SetWindowPos(NULL, rc.left, halfHeight - 21, 0, 0, SWP_NOSIZE);
			udtxtboxUWnd.GetWindowRect(&rc);
			ScreenToClient(&rc);
			udtxtboxUWnd.SetWindowPos(NULL, rc.left, halfHeight - 21, 0, 0, SWP_NOSIZE);

			txtbox0AWnd.SetWindowPos(NULL, 0, halfHeight + 5, remainingRectangle.Width(), halfHeight - 49, 0);
			txtbox1AWnd.SetWindowPos(NULL, 0, remainingRectangle.Height() - 44, remainingRectangle.Width(), 18, 0);
			ipaddrAWnd.SetWindowPos(NULL, 0, remainingRectangle.Height() - 21, 0, 0, SWP_NOSIZE);
			hkboxAWnd.GetWindowRect(&rc);
			ScreenToClient(&rc);
			hkboxAWnd.SetWindowPos(NULL, rc.left, remainingRectangle.Height() - 21, 0, 0, SWP_NOSIZE);
			udtxtboxAWnd.GetWindowRect(&rc);
			ScreenToClient(&rc);
			udtxtboxAWnd.SetWindowPos(NULL, rc.left, remainingRectangle.Height() - 21, 0, 0, SWP_NOSIZE);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.txtbox0U && GetDlgItem(IDC_TXTBOX0U).IsChild(hFocusedControl)) {
		controls.txtbox0U->About();
	} else if(controls.txtbox1U && GetDlgItem(IDC_TXTBOX1U).IsChild(hFocusedControl)) {
		controls.txtbox1U->About();
	} else if(controls.ipaddrU && GetDlgItem(IDC_IPADDRU).IsChild(hFocusedControl)) {
		controls.ipaddrU->About();
	} else if(controls.hkboxU && GetDlgItem(IDC_HKBOXU).IsChild(hFocusedControl)) {
		controls.hkboxU->About();
	} else if(controls.udtxtboxU && GetDlgItem(IDC_UDTXTBOXU).IsChild(hFocusedControl)) {
		controls.udtxtboxU->About();
	} else if(controls.txtbox0A && GetDlgItem(IDC_TXTBOX0A).IsChild(hFocusedControl)) {
		controls.txtbox0A->About();
	} else if(controls.txtbox1A && GetDlgItem(IDC_TXTBOX1A).IsChild(hFocusedControl)) {
		controls.txtbox1A->About();
	} else if(controls.ipaddrA && GetDlgItem(IDC_IPADDRA).IsChild(hFocusedControl)) {
		controls.ipaddrA->About();
	} else if(controls.hkboxA && GetDlgItem(IDC_HKBOXA).IsChild(hFocusedControl)) {
		controls.hkboxA->About();
	} else if(controls.udtxtboxA && GetDlgItem(IDC_UDTXTBOXA).IsChild(hFocusedControl)) {
		controls.udtxtboxA->About();
	}
	return 0;
}

LRESULT CMainDlg::OnSetFocusAbout(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled)
{
	hFocusedControl = reinterpret_cast<HWND>(wParam);
	bHandled = FALSE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void __stdcall CMainDlg::AbortedDragTxtbox0U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_AbortedDrag: index=0")));
}

void __stdcall CMainDlg::BeforeDrawTextTxtbox0U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_BeforeDrawText: index=0")));
}

void __stdcall CMainDlg::BeginDragTxtbox0U(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox0U->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxU_BeginDrag: index=0, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginRDragTxtbox0U(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox0U->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxU_BeginRDrag: index=0, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_Click: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTxtbox0U(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_ContextMenu: index=0, button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_DblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTxtbox0U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_DestroyedControlWindow: index=0, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveTxtbox0U(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_DragMouseMove: index=0, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_Drop: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownTxtbox0U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_KeyDown: index=0, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressTxtbox0U(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_KeyPress: index=0, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTxtbox0U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_KeyUp: index=0, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MDblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseDown: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseEnter: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseHover: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseLeave: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseMove: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseUp: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelTxtbox0U(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseWheel: index=0, button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLECompleteDrag: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLECompleteDrag: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				MessageBox(str, TEXT("Dragged files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragDrop: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragDrop: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.txtbox0U->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragEnter: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragEnter: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetTxtbox0U(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_OLEDragEnterPotentialTarget: index=0, hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveTxtbox0U(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragLeave: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragLeave: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetTxtbox0U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_OLEDragLeavePotentialTarget: index=0")));
}

void __stdcall CMainDlg::OLEDragMouseMoveTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragMouseMove: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragMouseMove: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackTxtbox0U(EditCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_OLEGiveFeedback: index=0, effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragTxtbox0U(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_OLEQueryContinueDrag: index=0, pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataTxtbox0U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEReceivedNewData: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEReceivedNewData: index=0, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataTxtbox0U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLESetData: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLESetData: index=0, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				LONG firstChar = 0;
				LONG lastChar = 0;
				controls.txtbox0U->GetDraggedTextRange(&firstChar, &lastChar);
				ATL::CString tmp = COLE2T(controls.txtbox0U->GetText());
				pData->SetData(formatID, _variant_t(tmp.Mid(firstChar, lastChar - firstChar + 1)), -1, 1);
				break;
			}
		}
	}
}

void __stdcall CMainDlg::OLEStartDragTxtbox0U(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEStartDrag: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEStartDrag: index=0, data=NULL");
	}

	AddLogEntry(str);

	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
	}
}

void __stdcall CMainDlg::OutOfMemoryTxtbox0U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_OutOfMemory: index=0")));
}

void __stdcall CMainDlg::RClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_RClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_RDblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTxtbox0U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_RecreatedControlWindow: index=0, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTxtbox0U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_ResizedControlWindow: index=0")));
}

void __stdcall CMainDlg::ScrollingTxtbox0U(EditCtlsLibU::ScrollAxisConstants axis)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_Scrolling: index=0, axis=%i"), axis);

	AddLogEntry(str);
}

void __stdcall CMainDlg::TextChangedTxtbox0U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_TextChanged: index=0")));
}

void __stdcall CMainDlg::TruncatedTextTxtbox0U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_TruncatedText: index=0")));
}

void __stdcall CMainDlg::WritingDirectionChangedTxtbox0U(EditCtlsLibU::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_WritingDirectionChanged: index=0, newWritingDirection=%i"), newWritingDirection);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_XClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTxtbox0U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_XDblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::AbortedDragTxtbox1U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_AbortedDrag: index=1")));
}

void __stdcall CMainDlg::BeforeDrawTextTxtbox1U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_BeforeDrawText: index=1")));
}

void __stdcall CMainDlg::BeginDragTxtbox1U(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox1U->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxU_BeginDrag: index=1, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginRDragTxtbox1U(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox1U->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxU_BeginRDrag: index=1, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_Click: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTxtbox1U(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_ContextMenu: index=1, button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_DblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTxtbox1U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_DestroyedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveTxtbox1U(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_DragMouseMove: index=1, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_Drop: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownTxtbox1U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_KeyDown: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressTxtbox1U(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_KeyPress: index=1, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTxtbox1U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_KeyUp: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseDown: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseEnter: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseHover: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseLeave: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseMove: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseUp: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelTxtbox1U(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_MouseWheel: index=1, button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLECompleteDrag: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLECompleteDrag: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				MessageBox(str, TEXT("Dragged files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragDrop: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragDrop: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.txtbox1U->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetTxtbox1U(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_OLEDragEnterPotentialTarget: index=1, hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragEnter: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragEnter: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetTxtbox1U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_OLEDragLeavePotentialTarget: index=1")));
}

void __stdcall CMainDlg::OLEDragLeaveTxtbox1U(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragLeave: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragLeave: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEDragMouseMove: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEDragMouseMove: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackTxtbox1U(EditCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_OLEGiveFeedback: index=1, effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragTxtbox1U(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_OLEQueryContinueDrag: index=1, pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataTxtbox1U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEReceivedNewData: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEReceivedNewData: index=1, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataTxtbox1U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLESetData: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLESetData: index=1, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				LONG firstChar = 0;
				LONG lastChar = 0;
				controls.txtbox1U->GetDraggedTextRange(&firstChar, &lastChar);
				ATL::CString tmp = COLE2T(controls.txtbox1U->GetText());
				pData->SetData(formatID, _variant_t(tmp.Mid(firstChar, lastChar - firstChar + 1)), -1, 1);
				break;
			}
		}
	}
}

void __stdcall CMainDlg::OLEStartDragTxtbox1U(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxU_OLEStartDrag: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxU_OLEStartDrag: index=1, data=NULL");
	}

	AddLogEntry(str);

	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
	}
}

void __stdcall CMainDlg::OutOfMemoryTxtbox1U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_OutOfMemory: index=1")));
}

void __stdcall CMainDlg::RClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_RClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_RDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTxtbox1U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_RecreatedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTxtbox1U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_ResizedControlWindow: index=1")));
}

void __stdcall CMainDlg::ScrollingTxtbox1U(EditCtlsLibU::ScrollAxisConstants axis)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_Scrolling: index=1, axis=%i"), axis);

	AddLogEntry(str);
}

void __stdcall CMainDlg::TextChangedTxtbox1U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_TextChanged: index=1")));
}

void __stdcall CMainDlg::TruncatedTextTxtbox1U()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxU_TruncatedText: index=1")));
}

void __stdcall CMainDlg::WritingDirectionChangedTxtbox1U(EditCtlsLibU::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_WritingDirectionChanged: index=1, newWritingDirection=%i"), newWritingDirection);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_XClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTxtbox1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxU_XDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::AddressChangedIpaddrU(long editBoxIndex, long* newFieldValue)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_AddressChanged: editBoxIndex=%i, newFieldValue=%i"), editBoxIndex, *newFieldValue);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_Click: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuIpaddrU(long editBoxIndex, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_ContextMenu: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), editBoxIndex, button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_DblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowIpaddrU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::FieldTextChangedIpaddrU(long editBoxIndex)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_FieldTextChanged: editBoxIndex=%i"), editBoxIndex);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownIpaddrU(long editBoxIndex, short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_KeyDown: editBoxIndex=%i, keyCode=%i, shift=%i"), editBoxIndex, *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressIpaddrU(long editBoxIndex, short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_KeyPress: editBoxIndex=%i, keyAscii=%i"), editBoxIndex, *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpIpaddrU(long editBoxIndex, short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_KeyUp: editBoxIndex=%i, keyCode=%i, shift=%i"), editBoxIndex, *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MDblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MouseDown: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MouseEnter: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MouseHover: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MouseLeave: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MouseMove: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MouseUp: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelIpaddrU(long editBoxIndex, short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_MouseWheel: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), editBoxIndex, button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropIpaddrU(long editBoxIndex, LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_OLEDragDrop: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.ipaddrU->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterIpaddrU(long editBoxIndex, LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_OLEDragEnter: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveIpaddrU(long editBoxIndex, LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_OLEDragLeave: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveIpaddrU(long editBoxIndex, LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_OLEDragMouseMove: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_RClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_RDblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowIpaddrU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowIpaddrU()
{
	AddLogEntry(CAtlString(TEXT("IPAddressU_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_XClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressU_XDblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowHkboxU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownHkboxU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressHkboxU(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpHkboxU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelHkboxU(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropHkboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.hkboxU->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterHkboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveHkboxU(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxU_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveHkboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowHkboxU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowHkboxU()
{
	AddLogEntry(CAtlString(TEXT("HKBoxU_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickHkboxU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxU_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeDrawTextUdtxtboxU()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxU_BeforeDrawText")));
}

void __stdcall CMainDlg::ChangedAcceleratorsUdtxtboxU(LPDISPATCH accelerators)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IUpDownAccelerators> pAccelerators = accelerators;
	if(pAccelerators) {
		str.Format(TEXT("UpDownTxtBoxU_ChangedAccelerators: accelerators=%i"), pAccelerators->Count());
	} else {
		str.Format(TEXT("UpDownTxtBoxU_ChangedAccelerators: accelerators=NULL"));
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingAcceleratorsUdtxtboxU(LPDISPATCH accelerators, VARIANT_BOOL* cancelChanges)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IUpDownAccelerators> pAccelerators = accelerators;
	if(pAccelerators) {
		str.Format(TEXT("UpDownTxtBoxU_ChangedAccelerators: accelerators=%i, cancelChanges=%i"), pAccelerators->Count(), *cancelChanges);
	} else {
		str.Format(TEXT("UpDownTxtBoxU_ChangedAccelerators: accelerators=NULL, cancelChanges=%i"), *cancelChanges);
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowUdtxtboxU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownUdtxtboxU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressUdtxtboxU(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpUdtxtboxU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_MouseWheel: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropUdtxtboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.udtxtboxU->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterUdtxtboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), *effect, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveUdtxtboxU(LPDISPATCH data, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxU_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveUdtxtboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), *effect, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryUdtxtboxU()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxU_OutOfMemory")));
}

void __stdcall CMainDlg::RClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowUdtxtboxU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowUdtxtboxU()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxU_ResizedControlWindow")));
}

void __stdcall CMainDlg::TextChangedUdtxtboxU()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxU_TextChanged")));
}

void __stdcall CMainDlg::TruncatedTextUdtxtboxU()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxU_TruncatedText")));
}

void __stdcall CMainDlg::ValueChangedUdtxtboxU()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxU_ValueChanged")));
}

void __stdcall CMainDlg::ValueChangingUdtxtboxU(long currentValue, long delta, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_ValueChanging: currentValue=%i, delta=%i, cancelChange=%i"), currentValue, delta, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::WritingDirectionChangedUdtxtboxU(EditCtlsLibU::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_WritingDirectionChanged: index=0, newWritingDirection=%i"), newWritingDirection);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxU_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::AbortedDragTxtbox0A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_AbortedDrag: index=0")));
}

void __stdcall CMainDlg::BeforeDrawTextTxtbox0A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_BeforeDrawText: index=0")));
}

void __stdcall CMainDlg::BeginDragTxtbox0A(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox0A->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxA_BeginDrag: index=0, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginRDragTxtbox0A(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox0A->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxA_BeginRDrag: index=0, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_Click: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTxtbox0A(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_ContextMenu: index=0, button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_DblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTxtbox0A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_DestroyedControlWindow: index=0, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveTxtbox0A(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_DragMouseMove: index=0, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_Drop: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownTxtbox0A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_KeyDown: index=0, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressTxtbox0A(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_KeyPress: index=0, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTxtbox0A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_KeyUp: index=0, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MDblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseDown: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseEnter: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseHover: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseLeave: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseMove: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseUp: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelTxtbox0A(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseWheel: index=0, button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLECompleteDrag: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLECompleteDrag: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				MessageBox(str, TEXT("Dragged files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragDrop: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragDrop: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.txtbox0A->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetTxtbox0A(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_OLEDragEnterPotentialTarget: index=0, hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragEnter: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragEnter: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetTxtbox0A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_OLEDragLeavePotentialTarget: index=0")));
}

void __stdcall CMainDlg::OLEDragLeaveTxtbox0A(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragLeave: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragLeave: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragMouseMove: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragMouseMove: index=0, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackTxtbox0A(EditCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_OLEGiveFeedback: index=0, effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragTxtbox0A(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_OLEQueryContinueDrag: index=0, pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataTxtbox0A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEReceivedNewData: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEReceivedNewData: index=0, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataTxtbox0A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLESetData: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLESetData: index=0, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				LONG firstChar = 0;
				LONG lastChar = 0;
				controls.txtbox0A->GetDraggedTextRange(&firstChar, &lastChar);
				ATL::CString tmp = COLE2T(controls.txtbox0A->GetText());
				pData->SetData(formatID, _variant_t(tmp.Mid(firstChar, lastChar - firstChar + 1)), -1, 1);
				break;
			}
		}
	}
}

void __stdcall CMainDlg::OLEStartDragTxtbox0A(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEStartDrag: index=0, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEStartDrag: index=0, data=NULL");
	}

	AddLogEntry(str);

	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
	}
}

void __stdcall CMainDlg::OutOfMemoryTxtbox0A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_OutOfMemory: index=0")));
}

void __stdcall CMainDlg::RClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_RClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_RDblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTxtbox0A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_RecreatedControlWindow: index=0, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTxtbox0A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_ResizedControlWindow: index=0")));
}

void __stdcall CMainDlg::ScrollingTxtbox0A(EditCtlsLibA::ScrollAxisConstants axis)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_Scrolling: index=0, axis=%i"), axis);

	AddLogEntry(str);
}

void __stdcall CMainDlg::TextChangedTxtbox0A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_TextChanged: index=0")));
}

void __stdcall CMainDlg::TruncatedTextTxtbox0A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_TruncatedText: index=0")));
}

void __stdcall CMainDlg::WritingDirectionChangedTxtbox0A(EditCtlsLibA::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_WritingDirectionChanged: index=0, newWritingDirection=%i"), newWritingDirection);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_XClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTxtbox0A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_XDblClick: index=0, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::AbortedDragTxtbox1A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_AbortedDrag: index=1")));
}

void __stdcall CMainDlg::BeforeDrawTextTxtbox1A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_BeforeDrawText: index=1")));
}

void __stdcall CMainDlg::BeginDragTxtbox1A(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox1A->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxA_BeginDrag: index=1, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginRDragTxtbox1A(long firstChar, long lastChar, short button, short shift, long x, long y)
{
	ATL::CString tmp = COLE2T(controls.txtbox1A->GetText());
	tmp.Mid(firstChar, lastChar - firstChar + 1);

	CAtlString str;
	str.Format(TEXT("TxtBoxA_BeginRDrag: index=1, firstChar=%i, lastChar=%i (%s), button=%i, shift=%i, x=%i, y=%i"), firstChar, lastChar, tmp, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_Click: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuTxtbox1A(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_ContextMenu: index=1, button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_DblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowTxtbox1A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_DestroyedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveTxtbox1A(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_DragMouseMove: index=1, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_Drop: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownTxtbox1A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_KeyDown: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressTxtbox1A(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_KeyPress: index=1, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpTxtbox1A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_KeyUp: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseDown: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseEnter: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseHover: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseLeave: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseMove: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseUp: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelTxtbox1A(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_MouseWheel: index=1, button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants performedEffect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLECompleteDrag: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLECompleteDrag: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				MessageBox(str, TEXT("Dragged files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragDrop: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragDrop: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.txtbox1A->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetTxtbox1A(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_OLEDragEnterPotentialTarget: index=1, hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragEnter: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragEnter: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetTxtbox1A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_OLEDragLeavePotentialTarget: index=1")));
}

void __stdcall CMainDlg::OLEDragLeaveTxtbox1A(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragLeave: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragLeave: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEDragMouseMove: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEDragMouseMove: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), *effect, button, shift, x, y, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackTxtbox1A(EditCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_OLEGiveFeedback: index=1, effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragTxtbox1A(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_OLEQueryContinueDrag: index=1, pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataTxtbox1A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEReceivedNewData: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEReceivedNewData: index=1, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataTxtbox1A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLESetData: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLESetData: index=1, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				LONG firstChar = 0;
				LONG lastChar = 0;
				controls.txtbox1A->GetDraggedTextRange(&firstChar, &lastChar);
				ATL::CString tmp = COLE2T(controls.txtbox1A->GetText());
				pData->SetData(formatID, _variant_t(tmp.Mid(firstChar, lastChar - firstChar + 1)), -1, 1);
				break;
			}
		}
	}
}

void __stdcall CMainDlg::OLEStartDragTxtbox1A(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("TxtBoxA_OLEStartDrag: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("TxtBoxA_OLEStartDrag: index=1, data=NULL");
	}

	AddLogEntry(str);

	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
	}
}

void __stdcall CMainDlg::OutOfMemoryTxtbox1A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_OutOfMemory: index=1")));
}

void __stdcall CMainDlg::RClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_RClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_RDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowTxtbox1A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_RecreatedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowTxtbox1A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_ResizedControlWindow: index=1")));
}

void __stdcall CMainDlg::ScrollingTxtbox1A(EditCtlsLibA::ScrollAxisConstants axis)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_Scrolling: index=1, axis=%i"), axis);

	AddLogEntry(str);
}

void __stdcall CMainDlg::TextChangedTxtbox1A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_TextChanged: index=1")));
}

void __stdcall CMainDlg::TruncatedTextTxtbox1A()
{
	AddLogEntry(CAtlString(TEXT("TxtBoxA_TruncatedText: index=1")));
}

void __stdcall CMainDlg::WritingDirectionChangedTxtbox1A(EditCtlsLibA::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_WritingDirectionChanged: index=1, newWritingDirection=%i"), newWritingDirection);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_XClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickTxtbox1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("TxtBoxA_XDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::AddressChangedIpaddrA(long editBoxIndex, long* newFieldValue)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_AddressChanged: editBoxIndex=%i, newFieldValue=%i"), editBoxIndex, *newFieldValue);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_Click: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuIpaddrA(long editBoxIndex, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_ContextMenu: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), editBoxIndex, button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_DblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowIpaddrA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::FieldTextChangedIpaddrA(long editBoxIndex)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_FieldTextChanged: editBoxIndex=%i"), editBoxIndex);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownIpaddrA(long editBoxIndex, short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_KeyDown: editBoxIndex=%i, keyCode=%i, shift=%i"), editBoxIndex, *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressIpaddrA(long editBoxIndex, short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_KeyPress: editBoxIndex=%i, keyAscii=%i"), editBoxIndex, *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpIpaddrA(long editBoxIndex, short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_KeyUp: editBoxIndex=%i, keyCode=%i, shift=%i"), editBoxIndex, *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MDblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MouseDown: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MouseEnter: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MouseHover: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MouseLeave: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MouseMove: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MouseUp: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelIpaddrA(long editBoxIndex, short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_MouseWheel: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), editBoxIndex, button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropIpaddrA(long editBoxIndex, LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_OLEDragDrop: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.ipaddrA->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterIpaddrA(long editBoxIndex, LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_OLEDragEnter: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveIpaddrA(long editBoxIndex, LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_OLEDragLeave: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveIpaddrA(long editBoxIndex, LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_OLEDragMouseMove: editBoxIndex=%i, data="), editBoxIndex);
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_RClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_RDblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowIpaddrA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowIpaddrA()
{
	AddLogEntry(CAtlString(TEXT("IPAddressA_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_XClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("IPAddressA_XDblClick: editBoxIndex=%i, button=%i, shift=%i, x=%i, y=%i"), editBoxIndex, button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowHkboxA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownHkboxA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressHkboxA(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpHkboxA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelHkboxA(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropHkboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.hkboxA->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterHkboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveHkboxA(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxA_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveHkboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("HKBoxA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("HKBoxA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowHkboxA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowHkboxA()
{
	AddLogEntry(CAtlString(TEXT("HKBoxA_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickHkboxA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("HKBoxA_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeDrawTextUdtxtboxA()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxA_BeforeDrawText")));
}

void __stdcall CMainDlg::ChangedAcceleratorsUdtxtboxA(LPDISPATCH accelerators)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IUpDownAccelerators> pAccelerators = accelerators;
	if(pAccelerators) {
		str.Format(TEXT("UpDownTxtBoxA_ChangedAccelerators: accelerators=%i"), pAccelerators->Count());
	} else {
		str.Format(TEXT("UpDownTxtBoxA_ChangedAccelerators: accelerators=NULL"));
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingAcceleratorsUdtxtboxA(LPDISPATCH accelerators, VARIANT_BOOL* cancelChanges)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IUpDownAccelerators> pAccelerators = accelerators;
	if(pAccelerators) {
		str.Format(TEXT("UpDownTxtBoxA_ChangedAccelerators: accelerators=%i, cancelChanges=%i"), pAccelerators->Count(), *cancelChanges);
	} else {
		str.Format(TEXT("UpDownTxtBoxA_ChangedAccelerators: accelerators=NULL, cancelChanges=%i"), *cancelChanges);
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_Click: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_ContextMenu: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_DblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowUdtxtboxA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownUdtxtboxA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressUdtxtboxA(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpUdtxtboxA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MouseDown: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MouseEnter: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MouseHover: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MouseLeave: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MouseMove: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MouseUp: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_MouseWheel: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropUdtxtboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.udtxtboxA->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterUdtxtboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), *effect, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveUdtxtboxA(LPDISPATCH data, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxA_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveUdtxtboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<EditCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("UpDownTxtBoxA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("UpDownTxtBoxA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), *effect, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OutOfMemoryUdtxtboxA()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxA_OutOfMemory")));
}

void __stdcall CMainDlg::RClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_RClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_RDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowUdtxtboxA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowUdtxtboxA()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxA_ResizedControlWindow")));
}

void __stdcall CMainDlg::TextChangedUdtxtboxA()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxA_TextChanged")));
}

void __stdcall CMainDlg::TruncatedTextUdtxtboxA()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxA_TruncatedText")));
}

void __stdcall CMainDlg::ValueChangedUdtxtboxA()
{
	AddLogEntry(CAtlString(TEXT("UpDownTxtBoxA_ValueChanged")));
}

void __stdcall CMainDlg::ValueChangingUdtxtboxA(long currentValue, long delta, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_ValueChanging: currentValue=%i, delta=%i, cancelChange=%i"), currentValue, delta, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::WritingDirectionChangedUdtxtboxA(EditCtlsLibA::WritingDirectionConstants newWritingDirection)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_WritingDirectionChanged: index=0, newWritingDirection=%i"), newWritingDirection);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_XClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("UpDownTxtBoxA_XDblClick: button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}
