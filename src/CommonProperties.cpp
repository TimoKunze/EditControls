// CommonProperties.cpp: The controls' "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));

	controls.invalidKeyCombinationsList.SubclassWindow(GetDlgItem(IDC_INVALIDKEYCOMBINATIONSBOX));
	hStateImageList = SetupStateImageList(controls.invalidKeyCombinationsList.GetImageList(LVSIL_STATE));
	controls.invalidKeyCombinationsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.invalidKeyCombinationsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.invalidKeyCombinationsList.AddColumn(TEXT(""), 0);

	controls.defaultModifierKeysList.SubclassWindow(GetDlgItem(IDC_DEFMODKEYSBOX));
	hStateImageList = SetupStateImageList(controls.defaultModifierKeysList.GetImageList(LVSIL_STATE));
	controls.defaultModifierKeysList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.defaultModifierKeysList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.defaultModifierKeysList.AddColumn(TEXT(""), 0);

	controls.currentValueEdit.Attach(GetDlgItem(IDC_CURRENTVALUEEDIT));
	controls.currentValueUpDown.Attach(GetDlgItem(IDC_CURRENTVALUEUPDOWN));

	// setup the toolbar
	WTL::CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	if(pNotificationDetails->hwndFrom == controls.disabledEventsList) {
		if(pDetails->iItem == properties.disabledEventsItemIndices.deMouseEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, MouseMove, MouseWheel"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deClickEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Click, DblClick, MClick, MDblClick, RClick, RDblClick, XClick, XDblClick"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deKeyboardEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("KeyDown, KeyUp, KeyPress"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deBeforeDrawText) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("BeforeDrawText"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deScrolling) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Scrolling"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deTextChangedEvents) {
			int numberOfIPAddressBoxes = 0;
			int numberOfTextBoxes = 0;
			for(UINT object = 0; object < m_nObjects; ++object) {
				LPUNKNOWN pControl = NULL;
				if(m_ppUnk[object]->QueryInterface(IID_IIPAddressBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					++numberOfIPAddressBoxes;
					pControl->Release();
				} else if(m_ppUnk[object]->QueryInterface(IID_ITextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					++numberOfTextBoxes;
					pControl->Release();
				} else if(m_ppUnk[object]->QueryInterface(IID_IUpDownTextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					++numberOfTextBoxes;
					pControl->Release();
				}
			}

			if(numberOfIPAddressBoxes == 0) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("TextChanged"))));
			} else if(numberOfTextBoxes == 0) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FieldTextChanged"))));
			} else {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FieldTextChanged, TextChanged"))));
			}
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deValueChangingEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ValueChanging, ValueChanged"))));
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	} else if(pNotificationDetails->hwndFrom == controls.invalidKeyCombinationsList) {
		switch(pDetails->iItem) {
			case 0:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations that no modifier key (Shift, Ctrl, Alt) is involved in, are invalid."))));
				break;
			case 1:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations of the form [SHIFT]+ are invalid."))));
				break;
			case 2:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations of the form [CTRL]+ are invalid."))));
				break;
			case 3:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations of the form [ALT]+ are invalid."))));
				break;
			case 4:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations of the form [SHIFT]+[CTRL]+ are invalid."))));
				break;
			case 5:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations of the form [SHIFT]+[ALT]+ are invalid."))));
				break;
			case 6:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations of the form [CTRL]+[ALT]+ are invalid."))));
				break;
			case 7:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Key combinations of the form [SHIFT]+[CTRL]+[ALT]+ are invalid."))));
				break;
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	} else if(pNotificationDetails->hwndFrom == controls.defaultModifierKeysList) {
		switch(pDetails->iItem) {
			case 0:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The default combination of modifier keys contains the [SHIFT] key."))));
				break;
			case 1:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The default combination of modifier keys contains the [CTRL] key."))));
				break;
			case 2:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The default combination of modifier keys contains the [ALT] key."))));
				break;
			case 3:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The default combination of modifier keys contains the extended key."))));
				break;
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));
	}

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IHotKeyBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IHotKeyBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IIPAddressBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IPAddressBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_ITextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ITextBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IUpDownTextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IUpDownTextBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IHotKeyBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("HotKeyBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IHotKeyBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IIPAddressBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("IPAddressBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IIPAddressBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_ITextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("TextBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ITextBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IUpDownTextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("UpDownTextBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IUpDownTextBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT CommonProperties::OnChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	SetDirty(TRUE);
	return 0;
}


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IHotKeyBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IHotKeyBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			reinterpret_cast<IHotKeyBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			InvalidKeyCombinationsConstants invalidKeyCombinations = static_cast<InvalidKeyCombinationsConstants>(0);
			reinterpret_cast<IHotKeyBox*>(pControl)->get_InvalidKeyCombinations(&invalidKeyCombinations);
			l = static_cast<LONG>(invalidKeyCombinations);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, ikcUnmodifiedKeys);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, ikcShiftPlusKey);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, ikcCtrlPlusKey);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, ikcAltPlusKey);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, ikcShiftPlusCtrlPlusKey);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, ikcShiftPlusAltPlusKey);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, ikcCtrlPlusAltPlusKey);
			SetBit(controls.invalidKeyCombinationsList.GetItemState(7, LVIS_STATEIMAGEMASK), l, ikcShiftPlusCtrlPlusAltPlusKey);
			reinterpret_cast<IHotKeyBox*>(pControl)->put_InvalidKeyCombinations(static_cast<InvalidKeyCombinationsConstants>(l));

			ModifierKeysConstants defaultModifierKeys = static_cast<ModifierKeysConstants>(0);
			reinterpret_cast<IHotKeyBox*>(pControl)->get_DefaultModifierKeys(&defaultModifierKeys);
			l = static_cast<LONG>(defaultModifierKeys);
			SetBit(controls.defaultModifierKeysList.GetItemState(0, LVIS_STATEIMAGEMASK), l, mkShift);
			SetBit(controls.defaultModifierKeysList.GetItemState(1, LVIS_STATEIMAGEMASK), l, mkCtrl);
			SetBit(controls.defaultModifierKeysList.GetItemState(2, LVIS_STATEIMAGEMASK), l, mkAlt);
			SetBit(controls.defaultModifierKeysList.GetItemState(3, LVIS_STATEIMAGEMASK), l, mkExt);
			reinterpret_cast<IHotKeyBox*>(pControl)->put_DefaultModifierKeys(static_cast<ModifierKeysConstants>(l));
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IIPAddressBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IIPAddressBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deTextChangedEvents, LVIS_STATEIMAGEMASK), l, deTextChangedEvents);
			reinterpret_cast<IIPAddressBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_ITextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<ITextBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deTextChangedEvents, LVIS_STATEIMAGEMASK), l, deTextChangedEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deBeforeDrawText, LVIS_STATEIMAGEMASK), l, deBeforeDrawText);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deValueChangingEvents, LVIS_STATEIMAGEMASK), l, deValueChangingEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deScrolling, LVIS_STATEIMAGEMASK), l, deScrolling);
			reinterpret_cast<ITextBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IUpDownTextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IUpDownTextBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deTextChangedEvents, LVIS_STATEIMAGEMASK), l, deTextChangedEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deBeforeDrawText, LVIS_STATEIMAGEMASK), l, deBeforeDrawText);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deValueChangingEvents, LVIS_STATEIMAGEMASK), l, deValueChangingEvents);
			reinterpret_cast<IUpDownTextBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			if(controls.currentValueUpDown.IsWindowEnabled()) {
				LONG oldValue = 0;
				LONG newValue = 0;
				reinterpret_cast<IUpDownTextBox*>(pControl)->get_CurrentValue(NULL, &oldValue);
				BOOL invalidValue = FALSE;
				newValue = controls.currentValueUpDown.GetPos32(&invalidValue);
				if(invalidValue) {
					controls.currentValueUpDown.SetPos32(oldValue);
					newValue = oldValue;
				}
				reinterpret_cast<IUpDownTextBox*>(pControl)->put_CurrentValue(NULL, newValue);
			}
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	int numberOfHotKeyBoxes = 0;
	int numberOfIPAddressBoxes = 0;
	int numberOfTextBoxes = 0;
	int numberOfUpDownTextBoxes = 0;
	DisabledEventsConstants* pDisabledEvents = reinterpret_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		ZeroMemory(pDisabledEvents, m_nObjects * sizeof(DisabledEventsConstants));
		InvalidKeyCombinationsConstants* pInvalidKeyCombinations = reinterpret_cast<InvalidKeyCombinationsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(InvalidKeyCombinationsConstants)));
		if(pInvalidKeyCombinations) {
			ZeroMemory(pInvalidKeyCombinations, m_nObjects * sizeof(InvalidKeyCombinationsConstants));
			ModifierKeysConstants* pDefaultModifierKeys = reinterpret_cast<ModifierKeysConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(ModifierKeysConstants)));
			if(pDefaultModifierKeys) {
				ZeroMemory(pDefaultModifierKeys, m_nObjects * sizeof(ModifierKeysConstants));
				for(UINT object = 0; object < m_nObjects; ++object) {
					LPUNKNOWN pControl = NULL;
					if(m_ppUnk[object]->QueryInterface(IID_IHotKeyBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
						++numberOfHotKeyBoxes;
						reinterpret_cast<IHotKeyBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
						reinterpret_cast<IHotKeyBox*>(pControl)->get_InvalidKeyCombinations(&pInvalidKeyCombinations[object]);
						reinterpret_cast<IHotKeyBox*>(pControl)->get_DefaultModifierKeys(&pDefaultModifierKeys[object]);
						pControl->Release();

					} else if(m_ppUnk[object]->QueryInterface(IID_IIPAddressBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
						++numberOfIPAddressBoxes;
						reinterpret_cast<IIPAddressBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
						pControl->Release();

					} else if(m_ppUnk[object]->QueryInterface(IID_ITextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
						++numberOfTextBoxes;
						reinterpret_cast<ITextBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
						pControl->Release();

					} else if(m_ppUnk[object]->QueryInterface(IID_IUpDownTextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
						++numberOfUpDownTextBoxes;
						reinterpret_cast<IUpDownTextBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
						pControl->Release();
					}
				}

				// fill the listboxes
				properties.disabledEventsItemIndices.Reset();
				LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
				properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
				controls.disabledEventsList.DeleteAllItems();
				properties.disabledEventsItemIndices.deMouseEvents = controls.disabledEventsList.AddItem(0, 0, TEXT("Mouse events"));
				controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deMouseEvents, CalculateStateImageMask(m_nObjects, pl, deMouseEvents), LVIS_STATEIMAGEMASK);
				properties.disabledEventsItemIndices.deClickEvents = controls.disabledEventsList.AddItem(1, 0, TEXT("Click events"));
				controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deClickEvents, CalculateStateImageMask(m_nObjects, pl, deClickEvents), LVIS_STATEIMAGEMASK);
				properties.disabledEventsItemIndices.deKeyboardEvents = controls.disabledEventsList.AddItem(2, 0, TEXT("Keyboard events"));
				controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, CalculateStateImageMask(m_nObjects, pl, deKeyboardEvents), LVIS_STATEIMAGEMASK);
				if(numberOfTextBoxes + numberOfUpDownTextBoxes > 0) {
					properties.disabledEventsItemIndices.deBeforeDrawText = controls.disabledEventsList.AddItem(3, 0, TEXT("BeforeDrawText event"));
					controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deBeforeDrawText, CalculateStateImageMask(m_nObjects, pl, deBeforeDrawText), LVIS_STATEIMAGEMASK);
				}
				if(numberOfTextBoxes > 0) {
					properties.disabledEventsItemIndices.deScrolling = controls.disabledEventsList.AddItem(4, 0, TEXT("Scrolling event"));
					controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deScrolling, CalculateStateImageMask(m_nObjects, pl, deScrolling), LVIS_STATEIMAGEMASK);
				}
				if(numberOfIPAddressBoxes + numberOfTextBoxes + numberOfUpDownTextBoxes > 0) {
					properties.disabledEventsItemIndices.deTextChangedEvents = controls.disabledEventsList.AddItem(5, 0, TEXT("Text changing events"));
					controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deTextChangedEvents, CalculateStateImageMask(m_nObjects, pl, deTextChangedEvents), LVIS_STATEIMAGEMASK);
				}
				if(numberOfUpDownTextBoxes > 0) {
					properties.disabledEventsItemIndices.deValueChangingEvents = controls.disabledEventsList.AddItem(6, 0, TEXT("Value changing events"));
					controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deValueChangingEvents, CalculateStateImageMask(m_nObjects, pl, deValueChangingEvents), LVIS_STATEIMAGEMASK);
				}
				controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				pl = reinterpret_cast<LONG*>(pInvalidKeyCombinations);
				properties.hWndCheckMarksAreSetFor = controls.invalidKeyCombinationsList;
				controls.invalidKeyCombinationsList.DeleteAllItems();
				if(numberOfHotKeyBoxes == 0) {
					GetDlgItem(IDC_INVALIDKEYCOMBINATIONSSTATIC).ShowWindow(SW_HIDE);
					controls.invalidKeyCombinationsList.ShowWindow(SW_HIDE);
				} else {
					GetDlgItem(IDC_INVALIDKEYCOMBINATIONSSTATIC).ShowWindow(SW_SHOW);
					controls.invalidKeyCombinationsList.ShowWindow(SW_SHOW);
					controls.invalidKeyCombinationsList.AddItem(0, 0, TEXT("Unmodified keys"));
					controls.invalidKeyCombinationsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, ikcUnmodifiedKeys), LVIS_STATEIMAGEMASK);
					controls.invalidKeyCombinationsList.AddItem(1, 0, TEXT("[SHIFT]+"));
					controls.invalidKeyCombinationsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, ikcShiftPlusKey), LVIS_STATEIMAGEMASK);
					controls.invalidKeyCombinationsList.AddItem(2, 0, TEXT("[CTRL]+"));
					controls.invalidKeyCombinationsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, ikcCtrlPlusKey), LVIS_STATEIMAGEMASK);
					controls.invalidKeyCombinationsList.AddItem(3, 0, TEXT("[ALT]+"));
					controls.invalidKeyCombinationsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, ikcAltPlusKey), LVIS_STATEIMAGEMASK);
					controls.invalidKeyCombinationsList.AddItem(4, 0, TEXT("[SHIFT]+[CTRL]+"));
					controls.invalidKeyCombinationsList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, ikcShiftPlusCtrlPlusKey), LVIS_STATEIMAGEMASK);
					controls.invalidKeyCombinationsList.AddItem(5, 0, TEXT("[SHIFT]+[ALT]+"));
					controls.invalidKeyCombinationsList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, ikcShiftPlusAltPlusKey), LVIS_STATEIMAGEMASK);
					controls.invalidKeyCombinationsList.AddItem(6, 0, TEXT("[CTRL]+[ALT]+"));
					controls.invalidKeyCombinationsList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, ikcCtrlPlusAltPlusKey), LVIS_STATEIMAGEMASK);
					controls.invalidKeyCombinationsList.AddItem(7, 0, TEXT("[SHIFT]+[CTRL]+[ALT]+"));
					controls.invalidKeyCombinationsList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, ikcShiftPlusCtrlPlusAltPlusKey), LVIS_STATEIMAGEMASK);
				}
				controls.invalidKeyCombinationsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				pl = reinterpret_cast<LONG*>(pDefaultModifierKeys);
				properties.hWndCheckMarksAreSetFor = controls.defaultModifierKeysList;
				controls.defaultModifierKeysList.DeleteAllItems();
				if(numberOfHotKeyBoxes == 0) {
					GetDlgItem(IDC_DEFMODKEYSSTATIC).ShowWindow(SW_HIDE);
					controls.defaultModifierKeysList.ShowWindow(SW_HIDE);
				} else {
					GetDlgItem(IDC_DEFMODKEYSSTATIC).ShowWindow(SW_SHOW);
					controls.defaultModifierKeysList.ShowWindow(SW_SHOW);
					controls.defaultModifierKeysList.AddItem(0, 0, TEXT("[SHIFT]"));
					controls.defaultModifierKeysList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, mkShift), LVIS_STATEIMAGEMASK);
					controls.defaultModifierKeysList.AddItem(1, 0, TEXT("[CTRL]"));
					controls.defaultModifierKeysList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, mkCtrl), LVIS_STATEIMAGEMASK);
					controls.defaultModifierKeysList.AddItem(2, 0, TEXT("[ALT]"));
					controls.defaultModifierKeysList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, mkAlt), LVIS_STATEIMAGEMASK);
					controls.defaultModifierKeysList.AddItem(3, 0, TEXT("Extended key"));
					controls.defaultModifierKeysList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, mkExt), LVIS_STATEIMAGEMASK);
				}
				controls.defaultModifierKeysList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				if(numberOfUpDownTextBoxes == 0) {
					GetDlgItem(IDC_CURRENTVALUESTATIC).ShowWindow(SW_HIDE);
					controls.currentValueEdit.ShowWindow(SW_HIDE);
					controls.currentValueUpDown.ShowWindow(SW_HIDE);

					GetDlgItem(IDC_INVALIDKEYCOMBINATIONSSTATIC).MoveWindow(186, 12, 130, 13);
					controls.invalidKeyCombinationsList.MoveWindow(186, 28, 175, 114);
				} else {
					GetDlgItem(IDC_CURRENTVALUESTATIC).ShowWindow(SW_SHOW);
					controls.currentValueEdit.ShowWindow(SW_SHOW);
					controls.currentValueUpDown.ShowWindow(SW_SHOW);

					GetDlgItem(IDC_INVALIDKEYCOMBINATIONSSTATIC).MoveWindow(186, 33, 130, 13);
					controls.invalidKeyCombinationsList.MoveWindow(186, 49, 175, 93);

					GetDlgItem(IDC_CURRENTVALUESTATIC).EnableWindow(numberOfUpDownTextBoxes == 1);
					controls.currentValueEdit.EnableWindow(numberOfUpDownTextBoxes == 1);
					controls.currentValueUpDown.EnableWindow(numberOfUpDownTextBoxes == 1);

					if(numberOfUpDownTextBoxes == 1) {
						for(UINT object = 0; object < m_nObjects; ++object) {
							LPUNKNOWN pControl = NULL;
							if(m_ppUnk[object]->QueryInterface(IID_IUpDownTextBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
								BaseConstants base = bDecimal;
								reinterpret_cast<IUpDownTextBox*>(pControl)->get_Base(&base);
								switch(base) {
									case bDecimal:
										controls.currentValueUpDown.SetBase(10);
										break;
									case bHexadecimal:
										controls.currentValueUpDown.SetBase(16);
										break;
								}
								LONG minimum = 0;
								reinterpret_cast<IUpDownTextBox*>(pControl)->get_Minimum(&minimum);
								LONG maximum = 0;
								reinterpret_cast<IUpDownTextBox*>(pControl)->get_Maximum(&maximum);
								controls.currentValueUpDown.SetRange32(minimum, maximum);
								LONG value = 0;
								reinterpret_cast<IUpDownTextBox*>(pControl)->get_CurrentValue(NULL, &value);
								controls.currentValueUpDown.SetPos32(value);

								pControl->Release();
								break;
							}
						}
					}
				}

				properties.hWndCheckMarksAreSetFor = NULL;

				HeapFree(GetProcessHeap(), 0, pDefaultModifierKeys);
			}
			HeapFree(GetProcessHeap(), 0, pInvalidKeyCombinations);
		}
		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}