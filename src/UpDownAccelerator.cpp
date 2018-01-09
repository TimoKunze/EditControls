// UpDownAccelerator.cpp: A wrapper for existing up down accelerators.

#include "stdafx.h"
#include "UpDownAccelerator.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP UpDownAccelerator::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IUpDownAccelerator, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


UpDownAccelerator::Properties::~Properties()
{
	if(pOwnerUpDownTextBox) {
		pOwnerUpDownTextBox->Release();
	}
}

HWND UpDownAccelerator::Properties::GetUpDownHWnd(void)
{
	ATLASSUME(pOwnerUpDownTextBox);

	OLE_HANDLE handle = NULL;
	pOwnerUpDownTextBox->get_hWndUpDown(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void UpDownAccelerator::Attach(int acceleratorIndex)
{
	properties.acceleratorIndex = acceleratorIndex;
}

void UpDownAccelerator::Detach(void)
{
	properties.acceleratorIndex = -1;
}

void UpDownAccelerator::SetOwner(UpDownTextBox* pOwner)
{
	if(properties.pOwnerUpDownTextBox) {
		properties.pOwnerUpDownTextBox->Release();
	}
	properties.pOwnerUpDownTextBox = pOwner;
	if(properties.pOwnerUpDownTextBox) {
		properties.pOwnerUpDownTextBox->AddRef();
	}
}


STDMETHODIMP UpDownAccelerator::get_ActivationTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	if(accelerators > 0) {
		LPUDACCEL pAccelerators = new UDACCEL[accelerators];
		SendMessage(hWndUpDown, UDM_GETACCEL, accelerators, reinterpret_cast<LPARAM>(pAccelerators));

		*pValue = pAccelerators[properties.acceleratorIndex].nSec;

		delete[] pAccelerators;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownAccelerator::put_ActivationTime(LONG newValue)
{
	HRESULT hr = E_FAIL;

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	if(accelerators > 0) {
		LPUDACCEL pAccelerators = new UDACCEL[accelerators];
		SendMessage(hWndUpDown, UDM_GETACCEL, accelerators, reinterpret_cast<LPARAM>(pAccelerators));

		pAccelerators[properties.acceleratorIndex].nSec = newValue;

		// sort...
		int newIndex = properties.acceleratorIndex;
		for(; (newIndex > 0) && (pAccelerators[properties.acceleratorIndex].nSec < pAccelerators[newIndex - 1].nSec); --newIndex);
		for(; (newIndex > 0) && (pAccelerators[properties.acceleratorIndex].nSec == pAccelerators[newIndex - 1].nSec) && (pAccelerators[newIndex].nInc < pAccelerators[newIndex - 1].nInc); --newIndex);
		if(newIndex == properties.acceleratorIndex) {
			for(; (newIndex < accelerators - 1) && (pAccelerators[properties.acceleratorIndex].nSec > pAccelerators[newIndex + 1].nSec); ++newIndex);
			for(; (newIndex < accelerators - 1) && (pAccelerators[properties.acceleratorIndex].nSec == pAccelerators[newIndex + 1].nSec) && (pAccelerators[properties.acceleratorIndex].nInc > pAccelerators[newIndex + 1].nInc); ++newIndex);
		}
		if(newIndex < properties.acceleratorIndex) {
			UDACCEL buffer = pAccelerators[properties.acceleratorIndex];
			CopyMemory(pAccelerators + newIndex + 1, pAccelerators + newIndex, (properties.acceleratorIndex - newIndex) * sizeof(UDACCEL));
			pAccelerators[newIndex] = buffer;
		} else if(newIndex > properties.acceleratorIndex) {
			UDACCEL buffer = pAccelerators[properties.acceleratorIndex];
			CopyMemory(pAccelerators + properties.acceleratorIndex, pAccelerators + properties.acceleratorIndex + 1, (newIndex - properties.acceleratorIndex) * sizeof(UDACCEL));
			pAccelerators[newIndex] = buffer;
		}
		properties.acceleratorIndex = newIndex;

		if(SendMessage(hWndUpDown, UDM_SETACCEL, accelerators, reinterpret_cast<LPARAM>(pAccelerators))) {
			hr = S_OK;
		}

		delete[] pAccelerators;
	}
	return hr;
}

STDMETHODIMP UpDownAccelerator::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.acceleratorIndex;
	return S_OK;
}

STDMETHODIMP UpDownAccelerator::get_StepSize(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	if(accelerators > 0) {
		LPUDACCEL pAccelerators = new UDACCEL[accelerators];
		SendMessage(hWndUpDown, UDM_GETACCEL, accelerators, reinterpret_cast<LPARAM>(pAccelerators));

		*pValue = pAccelerators[properties.acceleratorIndex].nInc;

		delete[] pAccelerators;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownAccelerator::put_StepSize(LONG newValue)
{
	HRESULT hr = E_FAIL;

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	if(accelerators > 0) {
		LPUDACCEL pAccelerators = new UDACCEL[accelerators];
		SendMessage(hWndUpDown, UDM_GETACCEL, accelerators, reinterpret_cast<LPARAM>(pAccelerators));

		pAccelerators[properties.acceleratorIndex].nInc = newValue;

		// sort...
		int newIndex = properties.acceleratorIndex;
		for(; (newIndex > 0) && (pAccelerators[properties.acceleratorIndex].nSec == pAccelerators[newIndex - 1].nSec) && (pAccelerators[newIndex].nInc < pAccelerators[newIndex - 1].nInc); --newIndex);
		if(newIndex == properties.acceleratorIndex) {
			for(; (newIndex < accelerators - 1) && (pAccelerators[properties.acceleratorIndex].nSec == pAccelerators[newIndex + 1].nSec) && (pAccelerators[properties.acceleratorIndex].nInc > pAccelerators[newIndex + 1].nInc); ++newIndex);
		}
		if(newIndex < properties.acceleratorIndex) {
			UDACCEL buffer = pAccelerators[properties.acceleratorIndex];
			CopyMemory(pAccelerators + newIndex + 1, pAccelerators + newIndex, (properties.acceleratorIndex - newIndex) * sizeof(UDACCEL));
			pAccelerators[newIndex] = buffer;
		} else if(newIndex > properties.acceleratorIndex) {
			UDACCEL buffer = pAccelerators[properties.acceleratorIndex];
			CopyMemory(pAccelerators + properties.acceleratorIndex, pAccelerators + properties.acceleratorIndex + 1, (newIndex - properties.acceleratorIndex) * sizeof(UDACCEL));
			pAccelerators[newIndex] = buffer;
		}
		properties.acceleratorIndex = newIndex;

		if(SendMessage(hWndUpDown, UDM_SETACCEL, accelerators, reinterpret_cast<LPARAM>(pAccelerators))) {
			hr = S_OK;
		}

		delete[] pAccelerators;
	}
	return hr;
}