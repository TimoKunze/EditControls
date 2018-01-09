// UpDownAccelerators.cpp: Manages a collection of UpDownAccelerator objects

#include "stdafx.h"
#include "UpDownAccelerators.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP UpDownAccelerators::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IUpDownAccelerators, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP UpDownAccelerators::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<UpDownAccelerators>* pUpDownAccelsObj = NULL;
	CComObject<UpDownAccelerators>::CreateInstance(&pUpDownAccelsObj);
	pUpDownAccelsObj->AddRef();

	// clone all settings
	pUpDownAccelsObj->properties = properties;
	if(pUpDownAccelsObj->properties.pOwnerUpDownTextBox) {
		pUpDownAccelsObj->properties.pOwnerUpDownTextBox->AddRef();
	}

	pUpDownAccelsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pUpDownAccelsObj->Release();
	return S_OK;
}

STDMETHODIMP UpDownAccelerators::Next(ULONG numberOfMaxAccelerators, VARIANT* pAccelerators, ULONG* pNumberOfAcceleratorsReturned)
{
	ATLASSERT_POINTER(pAccelerators, VARIANT);
	if(!pAccelerators) {
		return E_POINTER;
	}

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));

	ULONG i = 0;
	for(i = 0; i < numberOfMaxAccelerators; ++i) {
		VariantInit(&pAccelerators[i]);

		if(properties.nextEnumeratedAccelerator >= accelerators) {
			// there's nothing more to iterate
			break;
		}

		ClassFactory::InitAccelerator(properties.nextEnumeratedAccelerator, properties.pOwnerUpDownTextBox, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pAccelerators[i].pdispVal), FALSE);
		pAccelerators[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedAccelerator;
	}
	if(pNumberOfAcceleratorsReturned) {
		*pNumberOfAcceleratorsReturned = i;
	}

	return (i == numberOfMaxAccelerators ? S_OK : S_FALSE);
}

STDMETHODIMP UpDownAccelerators::Reset(void)
{
	properties.nextEnumeratedAccelerator = 0;
	return S_OK;
}

STDMETHODIMP UpDownAccelerators::Skip(ULONG numberOfAcceleratorsToSkip)
{
	properties.nextEnumeratedAccelerator += static_cast<int>(numberOfAcceleratorsToSkip);
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


UpDownAccelerators::Properties::~Properties()
{
	if(pOwnerUpDownTextBox) {
		pOwnerUpDownTextBox->Release();
	}
}

HWND UpDownAccelerators::Properties::GetUpDownHWnd(void)
{
	ATLASSUME(pOwnerUpDownTextBox);

	OLE_HANDLE handle = NULL;
	pOwnerUpDownTextBox->get_hWndUpDown(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void UpDownAccelerators::SetOwner(UpDownTextBox* pOwner)
{
	if(properties.pOwnerUpDownTextBox) {
		properties.pOwnerUpDownTextBox->Release();
	}
	properties.pOwnerUpDownTextBox = pOwner;
	if(properties.pOwnerUpDownTextBox) {
		properties.pOwnerUpDownTextBox->AddRef();
	}
}


STDMETHODIMP UpDownAccelerators::get_Item(LONG acceleratorIndex, IUpDownAccelerator** ppAccelerator)
{
	ATLASSERT_POINTER(ppAccelerator, IUpDownAccelerator*);
	if(!ppAccelerator) {
		return E_POINTER;
	}

	if(IsValidAcceleratorIndex(acceleratorIndex, properties.pOwnerUpDownTextBox)) {
		ClassFactory::InitAccelerator(acceleratorIndex, properties.pOwnerUpDownTextBox, IID_IUpDownAccelerator, reinterpret_cast<LPUNKNOWN*>(ppAccelerator), FALSE);
		return S_OK;
	}
	return DISP_E_BADINDEX;
}

STDMETHODIMP UpDownAccelerators::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP UpDownAccelerators::Add(LONG activationTime, LONG stepSize, IUpDownAccelerator** ppAddedAccelerator/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedAccelerator, IUpDownAccelerator*);
	if(!ppAddedAccelerator) {
		return E_POINTER;
	}

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	int insertAt = 0;
	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	LPUDACCEL pAccelerators = reinterpret_cast<LPUDACCEL>(HeapAlloc(GetProcessHeap(), 0, (accelerators + 1) * sizeof(UDACCEL)));
	if(!pAccelerators) {
		return E_OUTOFMEMORY;
	}
	if(accelerators > 0) {
		LPUDACCEL pCurrentAccelerators = reinterpret_cast<LPUDACCEL>(HeapAlloc(GetProcessHeap(), 0, accelerators * sizeof(UDACCEL)));
		if(!pCurrentAccelerators) {
			HeapFree(GetProcessHeap(), 0, pAccelerators);
			return E_OUTOFMEMORY;
		}
		SendMessage(hWndUpDown, UDM_GETACCEL, accelerators, reinterpret_cast<LPARAM>(pCurrentAccelerators));

		for(insertAt = 0; insertAt < accelerators; ++insertAt) {
			if(static_cast<UINT>(activationTime) < pCurrentAccelerators[insertAt].nSec) {
				break;
			} else if(static_cast<UINT>(activationTime) == pCurrentAccelerators[insertAt].nSec) {
				// sort by step size (smaller one wins)
				while((static_cast<UINT>(activationTime) == pCurrentAccelerators[insertAt].nSec) && (pCurrentAccelerators[insertAt].nInc < static_cast<UINT>(stepSize))) {
					if(++insertAt == accelerators) {
						break;
					}
				}
				break;
			}
		}

		if(insertAt > 0) {
			CopyMemory(pAccelerators, pCurrentAccelerators, insertAt * sizeof(UDACCEL));
		}
		if(insertAt < accelerators) {
			CopyMemory(pAccelerators + insertAt + 1, pCurrentAccelerators + insertAt, (accelerators - insertAt) * sizeof(UDACCEL));
		}
		HeapFree(GetProcessHeap(), 0, pCurrentAccelerators);
	}
	pAccelerators[insertAt].nSec = static_cast<UINT>(activationTime);
	pAccelerators[insertAt].nInc = static_cast<UINT>(stepSize);

	// finally insert the accelerator
	HRESULT hr = E_FAIL;
	if(SendMessage(hWndUpDown, UDM_SETACCEL, accelerators + 1, reinterpret_cast<LPARAM>(pAccelerators))) {
		ClassFactory::InitAccelerator(insertAt, properties.pOwnerUpDownTextBox, IID_IUpDownAccelerator, reinterpret_cast<LPUNKNOWN*>(ppAddedAccelerator));
		hr = S_OK;
	}

	HeapFree(GetProcessHeap(), 0, pAccelerators);
	return hr;
}

STDMETHODIMP UpDownAccelerators::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	*pValue = static_cast<LONG>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	return S_OK;
}

STDMETHODIMP UpDownAccelerators::Remove(LONG acceleratorIndex)
{
	HRESULT hr = DISP_E_BADINDEX;

	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	if((accelerators >= 0) && (acceleratorIndex < accelerators)) {
		hr = E_FAIL;

		if(accelerators == 1) {
			UDACCEL accel = {0};
			if(SendMessage(hWndUpDown, UDM_SETACCEL, 1, reinterpret_cast<LPARAM>(&accel))) {
				hr = S_OK;
			}
		} else {
			LPUDACCEL pAccelerators = new UDACCEL[accelerators];
			SendMessage(hWndUpDown, UDM_GETACCEL, accelerators, reinterpret_cast<LPARAM>(pAccelerators));

			if(acceleratorIndex < accelerators - 1) {
				CopyMemory(pAccelerators + acceleratorIndex, pAccelerators + acceleratorIndex + 1, (accelerators - acceleratorIndex - 1) * sizeof(UDACCEL));
			}

			if(SendMessage(hWndUpDown, UDM_SETACCEL, accelerators - 1, reinterpret_cast<LPARAM>(pAccelerators))) {
				hr = S_OK;
			}

			delete[] pAccelerators;
		}
	}
	return hr;
}

STDMETHODIMP UpDownAccelerators::RemoveAll(void)
{
	HWND hWndUpDown = properties.GetUpDownHWnd();
	ATLASSERT(IsWindow(hWndUpDown));

	UDACCEL accel = {0};
	if(SendMessage(hWndUpDown, UDM_SETACCEL, 1, reinterpret_cast<LPARAM>(&accel))) {
		return S_OK;
	}
	return E_FAIL;
}