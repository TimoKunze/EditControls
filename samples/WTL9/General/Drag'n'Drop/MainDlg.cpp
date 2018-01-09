// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

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
	if(controls.txtboxU) {
		DispEventUnadvise(controls.txtboxU);
		controls.txtboxU.Release();
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

	controls.oleDDCheck = GetDlgItem(IDC_OLEDDCHECK);
	controls.oleDDCheck.SetCheck(BST_CHECKED);
	txtBoxUWnd.SubclassWindow(GetDlgItem(IDC_TXTBOXU));
	txtBoxUWnd.QueryControl(__uuidof(ITextBox), reinterpret_cast<LPVOID*>(&controls.txtboxU));
	if(controls.txtboxU) {
		DispEventAdvise(controls.txtboxU);
		LoadText();
	}

	return TRUE;
}

LRESULT CMainDlg::OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(wParam == TIMERID_CONTEXTMENUHACK) {
		KillTimer(TIMERID_CONTEXTMENUHACK);
		suppressDefaultContextMenu = FALSE;
	} else {
		bHandled = FALSE;
	}
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.txtboxU) {
		controls.txtboxU->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::LoadText(void)
{
	HANDLE hFile = CreateFile(TEXT("MainDlg.cpp"), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if(hFile != INVALID_HANDLE_VALUE) {
		DWORD bufferSize = GetFileSize(hFile, NULL);
		CHAR* pBuffer = new CHAR[bufferSize + 1];
		memset(pBuffer, 0, bufferSize + 1);
		DWORD bytesRead = 0;
		ReadFile(hFile, pBuffer, bufferSize, &bytesRead, NULL);

		controls.txtboxU->PutText(CComBSTR(pBuffer).Detach());

		delete[] pBuffer,
		CloseHandle(hFile);
	}
}

void __stdcall CMainDlg::AbortedDragTxtboxu()
{
	controls.txtboxU->SetInsertMarkPosition(impNowhere, -1);
}

void __stdcall CMainDlg::BeginDragTxtboxu(LONG firstChar, LONG lastChar, short /*button*/, short /*shift*/, long /*x*/, long /*y*/)
{
	bRightDrag = FALSE;
	if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
		IOLEDataObject* p = NULL;
		controls.txtboxU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), firstChar, lastChar, 0);
	} else {
		controls.txtboxU->BeginDrag(firstChar, lastChar, static_cast<OLE_HANDLE>(-1), NULL, NULL);
	}
}

void __stdcall CMainDlg::BeginRDragTxtboxu(LONG firstChar, LONG lastChar, short /*button*/, short /*shift*/, long /*x*/, long /*y*/)
{
	bRightDrag = TRUE;
	if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
		IOLEDataObject* p = NULL;
		controls.txtboxU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), firstChar, lastChar, 0);
	} else {
		controls.txtboxU->BeginDrag(firstChar, lastChar, static_cast<OLE_HANDLE>(-1), NULL, NULL);
	}
}

void __stdcall CMainDlg::ContextMenuTxtboxu(short /*button*/, short /*shift*/, long /*x*/, long /*y*/, VARIANT_BOOL* showDefaultMenu)
{
	*showDefaultMenu = (suppressDefaultContextMenu ? VARIANT_FALSE : VARIANT_TRUE);
}

void __stdcall CMainDlg::DragMouseMoveTxtboxu(short /*button*/, short /*shift*/, long x, long y, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	LONG newCharIndex = 0;
	InsertMarkPositionConstants newInsertMarkRelativePosition = impNowhere;
	controls.txtboxU->GetClosestInsertMarkPosition(x, y, &newInsertMarkRelativePosition, &newCharIndex);
	controls.txtboxU->SetInsertMarkPosition(newInsertMarkRelativePosition, newCharIndex);
}

void __stdcall CMainDlg::DropTxtboxu(short /*button*/, short /*shift*/, long x, long y)
{
	if(bRightDrag) {
		HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_DRAGMENU));
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.txtboxU->GethWnd())), &pt);
		UINT selectedCmd = TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.txtboxU->GethWnd())), NULL);
		DestroyMenu(hPopupMenu);
		DestroyMenu(hMenu);

		switch(selectedCmd) {
			case ID_ACTIONS_CANCEL:
				AbortedDragTxtboxu();
				return;
				break;
			case ID_ACTIONS_COPYTEXT: {
				InsertMarkPositionConstants insertionMark = impNowhere;
				LONG insertAt = 0;
				controls.txtboxU->GetInsertMarkPosition(&insertionMark, &insertAt);
				if(insertionMark != impNowhere) {
					LONG firstDraggedChar = 0;
					LONG lastDraggedChar = 0;
					controls.txtboxU->GetDraggedTextRange(&firstDraggedChar, &lastDraggedChar);
					ATL::CString oldText = COLE2T(controls.txtboxU->GetText());
					CComBSTR draggedText = oldText.Mid(firstDraggedChar, lastDraggedChar - firstDraggedChar + 1);

					if(insertionMark == impAfter) {
						++insertAt;
					}

					CComBSTR newText = oldText.Left(insertAt);
					newText.AppendBSTR(draggedText);
					newText.Append(oldText.Mid(insertAt));

					controls.txtboxU->PutText(newText.Detach());
					controls.txtboxU->SetSelection(insertAt, insertAt + (lastDraggedChar - firstDraggedChar + 1));
				}

				controls.txtboxU->SetInsertMarkPosition(impNowhere, -1);
				return;
				break;
			}
			case ID_ACTIONS_MOVETEXT:
				// fall through
				break;
		}
	}

	// move the text
	InsertMarkPositionConstants insertionMark = impNowhere;
	LONG insertAt = 0;
	controls.txtboxU->GetInsertMarkPosition(&insertionMark, &insertAt);
	if(insertionMark != impNowhere) {
		LONG firstDraggedChar = 0;
		LONG lastDraggedChar = 0;
		controls.txtboxU->GetDraggedTextRange(&firstDraggedChar, &lastDraggedChar);
		ATL::CString oldText = COLE2T(controls.txtboxU->GetText());
		CComBSTR draggedText = oldText.Mid(firstDraggedChar, lastDraggedChar - firstDraggedChar + 1);

		if(insertionMark == impAfter) {
			++insertAt;
		}

		if(insertAt < firstDraggedChar) {
			CComBSTR newText = L"";
			if(insertAt > 0) {
				newText = oldText.Left(insertAt);
			}
			newText.AppendBSTR(draggedText);
			newText.Append(oldText.Mid(insertAt, firstDraggedChar - insertAt));
			newText.Append(oldText.Mid(lastDraggedChar + 1));

			controls.txtboxU->PutText(newText.Detach());
			controls.txtboxU->SetSelection(insertAt, insertAt + (lastDraggedChar - firstDraggedChar + 1));
		} else if(insertAt > (lastDraggedChar + 1)) {
			CComBSTR newText = L"";
			if(firstDraggedChar > 0) {
				newText = oldText.Left(firstDraggedChar);
			}
			newText.Append(oldText.Mid(lastDraggedChar + 1, insertAt - lastDraggedChar - 1));
			newText.AppendBSTR(draggedText);
			newText.Append(oldText.Mid(insertAt));

			controls.txtboxU->PutText(newText.Detach());
			controls.txtboxU->SetSelection(firstDraggedChar + (insertAt - lastDraggedChar - 1), insertAt);
		} else {
			// restore the selection
			controls.txtboxU->SetSelection(firstDraggedChar, lastDraggedChar + 1);
		}
	}
	controls.txtboxU->SetInsertMarkPosition(impNowhere, -1);
}

void __stdcall CMainDlg::KeyDownTxtboxu(short* keyCode, short /*shift*/)
{
	if(*keyCode == VK_ESCAPE) {
		LONG firstChar = 0;
		LONG lastChar = 0;
		controls.txtboxU->GetDraggedTextRange(&firstChar, &lastChar);
		if((firstChar >= 0) && (lastChar >= 0)) {
			controls.txtboxU->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::MouseUpTxtboxu(short button, short /*shift*/, long /*x*/, long /*y*/)
{
	if(controls.oleDDCheck.GetCheck() == BST_UNCHECKED) {
		if(((button == 1) && !bRightDrag) || ((button == 2) && bRightDrag)) {
			LONG firstChar = 0;
			LONG lastChar = 0;
			controls.txtboxU->GetDraggedTextRange(&firstChar, &lastChar);
			if((firstChar >= 0) && (lastChar >= 0)) {
				InsertMarkPositionConstants insertionMark = impNowhere;
				LONG insertAt = 0;
				controls.txtboxU->GetInsertMarkPosition(&insertionMark, &insertAt);
				suppressDefaultContextMenu = TRUE;
				if(insertionMark == impNowhere) {
					// cancel
					controls.txtboxU->EndDrag(VARIANT_TRUE);
				} else {
					// drop
					controls.txtboxU->EndDrag(VARIANT_FALSE);
				}
				// suppressing the default context menu is a bit tricky - use a timer to reset the flag
				SetTimer(TIMERID_CONTEXTMENUHACK, 50);
			}
		}
	}
}

void __stdcall CMainDlg::OLECompleteDragTxtboxu(LPDISPATCH data, long performedEffect)
{
	CComQIPtr<IDataObject> pData = data;
	if(performedEffect == odeMove) {
		// remove the dragged text
		LONG firstChar = 0;
		LONG lastChar = 0;
		controls.txtboxU->GetDraggedTextRange(&firstChar, &lastChar);
		controls.txtboxU->SetSelection(firstChar, lastChar + 1);
		controls.txtboxU->ReplaceSelectedText(_bstr_t(TEXT("")).Detach(), VARIANT_FALSE);
	} else if(pData) {
		FORMATETC format = {CF_TARGETCLSID, NULL, 1, -1, TYMED_HGLOBAL};
		if(pData->QueryGetData(&format) == S_OK) {
			STGMEDIUM storageMedium = {0};
			if(pData->GetData(&format, &storageMedium) == S_OK) {
				GUID* pCLSIDOfTarget = (GUID*) GlobalLock(storageMedium.hGlobal);
				if(*pCLSIDOfTarget == CLSID_RecycleBin) {
					// remove the dragged text
					LONG firstChar = 0;
					LONG lastChar = 0;
					controls.txtboxU->GetDraggedTextRange(&firstChar, &lastChar);
					controls.txtboxU->SetSelection(firstChar, lastChar + 1);
					controls.txtboxU->ReplaceSelectedText(_bstr_t(TEXT("")).Detach(), VARIANT_FALSE);
				}
				GlobalUnlock(storageMedium.hGlobal);
				ReleaseStgMedium(&storageMedium);
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropTxtboxu(LPDISPATCH data, long* effect, short /*button*/, short shift, long /*x*/, long /*y*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		CComBSTR str = NULL;
		if(pData->GetFormat(CF_UNICODETEXT, -1, 1) != VARIANT_FALSE) {
			str = pData->GetData(CF_UNICODETEXT, -1, 1).bstrVal;
		} else if(pData->GetFormat(CF_TEXT, -1, 1) != VARIANT_FALSE) {
			str = pData->GetData(CF_TEXT, -1, 1).bstrVal;
		} else if(pData->GetFormat(CF_OEMTEXT, -1, 1) != VARIANT_FALSE) {
			str = pData->GetData(CF_OEMTEXT, -1, 1).bstrVal;
		} else if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			// insert a line for each file/folder
			_variant_t v = pData->GetData(CF_HDROP, -1, 1);
			CComSafeArray<BSTR> arr;
			arr.Attach(v.parray);
			for(LONG i = arr.GetLowerBound(); i <= arr.GetUpperBound(); i++) {
				ATL::CString tmp = COLE2T(arr.GetAt(i));
				int p = tmp.ReverseFind(TEXT('\\'));
				if(p >= 0) {
					str.Append(tmp.Mid(p + 1));
				} else {
					str.Append(tmp);
				}
			}
			arr.Detach();
		}

		if(str != TEXT("")) {
			// insert the text
			InsertMarkPositionConstants insertionMark;
			LONG insertAt = 0;
			controls.txtboxU->GetInsertMarkPosition(&insertionMark, &insertAt);

			switch(insertionMark) {
				case impAfter:
					controls.txtboxU->SetSelection(insertAt + 1, insertAt + 1);
					controls.txtboxU->ReplaceSelectedText(str.Detach(), VARIANT_FALSE);
					break;
				case impBefore:
					controls.txtboxU->SetSelection(insertAt, insertAt);
					controls.txtboxU->ReplaceSelectedText(str.Detach(), VARIANT_FALSE);
					break;
			}
		}
	}

	controls.txtboxU->SetInsertMarkPosition(impNowhere, -1);

	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}
}

void __stdcall CMainDlg::OLEDragEnterTxtboxu(LPDISPATCH /*data*/, long* effect, short /*button*/, short shift, long x, long y, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	LONG newCharIndex = 0;
	InsertMarkPositionConstants newInsertMarkRelativePosition = impNowhere;
	controls.txtboxU->GetClosestInsertMarkPosition(x, y, &newInsertMarkRelativePosition, &newCharIndex);
	controls.txtboxU->SetInsertMarkPosition(newInsertMarkRelativePosition, newCharIndex);

	*effect = odeMove;
	if(shift & 1/*vbShiftMask*/) {
		*effect = odeMove;
	}
	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}
}

void __stdcall CMainDlg::OLEDragLeaveTxtboxu(LPDISPATCH /*data*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/)
{
	controls.txtboxU->SetInsertMarkPosition(impNowhere, -1);
}

void __stdcall CMainDlg::OLEDragMouseMoveTxtboxu(LPDISPATCH /*data*/, long* effect, short /*button*/, short shift, long x, long y, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	LONG newCharIndex = 0;
	InsertMarkPositionConstants newInsertMarkRelativePosition = impNowhere;
	controls.txtboxU->GetClosestInsertMarkPosition(x, y, &newInsertMarkRelativePosition, &newCharIndex);
	controls.txtboxU->SetInsertMarkPosition(newInsertMarkRelativePosition, newCharIndex);

	*effect = odeMove;
	if(shift & 1/*vbShiftMask*/) {
		*effect = odeMove;
	}
	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}
}

void __stdcall CMainDlg::OLESetDataTxtboxu(LPDISPATCH data, long formatID, long /*index*/, long /*dataOrViewAspect*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				LONG firstChar = 0;
				LONG lastChar = 0;
				controls.txtboxU->GetDraggedTextRange(&firstChar, &lastChar);
				ATL::CString tmp = COLE2T(controls.txtboxU->GetText());
				pData->SetData(formatID, _variant_t(tmp.Mid(firstChar, lastChar - firstChar + 1)), -1, 1);
				break;
			}
		}
	}
}

void __stdcall CMainDlg::OLEStartDragTxtboxu(LPDISPATCH data)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
	}
}

void __stdcall CMainDlg::RecreatedControlWindowTxtboxu(long /*hWnd*/)
{
	LoadText();
}