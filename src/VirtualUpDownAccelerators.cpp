// VirtualUpDownAccelerators.cpp: Manages a collection of VirtualUpDownAccelerator objects

#include "stdafx.h"
#include "VirtualUpDownAccelerators.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualUpDownAccelerators::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualUpDownAccelerators, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP VirtualUpDownAccelerators::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<VirtualUpDownAccelerators>* pVUpDownAccelsObj = NULL;
	CComObject<VirtualUpDownAccelerators>::CreateInstance(&pVUpDownAccelsObj);
	pVUpDownAccelsObj->AddRef();

	// clone all settings
	pVUpDownAccelsObj->properties = properties;

	pVUpDownAccelsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pVUpDownAccelsObj->Release();
	return S_OK;
}

STDMETHODIMP VirtualUpDownAccelerators::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		if(properties.nextEnumeratedAccelerator >= properties.numberOfAccelerators) {
			// there's nothing more to iterate
			break;
		}

		pItems[i].pdispVal = NULL;

		// create and initialize a VirtualUpDownAccelerator object
		CComObject<VirtualUpDownAccelerator>* pVUpDownAccelObj = NULL;
		CComObject<VirtualUpDownAccelerator>::CreateInstance(&pVUpDownAccelObj);
		pVUpDownAccelObj->AddRef();
		pVUpDownAccelObj->Attach(properties.pAccelerators, properties.nextEnumeratedAccelerator);
		pVUpDownAccelObj->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&pItems[i].pdispVal));
		pVUpDownAccelObj->Release();

		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedAccelerator;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP VirtualUpDownAccelerators::Reset(void)
{
	properties.nextEnumeratedAccelerator = 0;
	return S_OK;
}

STDMETHODIMP VirtualUpDownAccelerators::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedAccelerator += static_cast<int>(numberOfItemsToSkip);
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


void VirtualUpDownAccelerators::Attach(int numberOfAccelerators, LPUDACCEL pAccelerators)
{
	properties.numberOfAccelerators = numberOfAccelerators;
	properties.pAccelerators = pAccelerators;
}

void VirtualUpDownAccelerators::Detach(void)
{
	properties.numberOfAccelerators = 0;
	if(properties.pAccelerators) {
		delete[] properties.pAccelerators;
		properties.pAccelerators = NULL;
	}
}


STDMETHODIMP VirtualUpDownAccelerators::get_Item(LONG acceleratorIndex, IVirtualUpDownAccelerator** ppAccelerator)
{
	ATLASSERT_POINTER(ppAccelerator, IVirtualUpDownAccelerator*);
	if(!ppAccelerator) {
		return E_POINTER;
	}

	if((acceleratorIndex >= 0) && (acceleratorIndex < properties.numberOfAccelerators)) {
		*ppAccelerator = NULL;

		// create and initialize a VirtualUpDownAccelerator object
		CComObject<VirtualUpDownAccelerator>* pVUpDownAccelObj = NULL;
		CComObject<VirtualUpDownAccelerator>::CreateInstance(&pVUpDownAccelObj);
		pVUpDownAccelObj->AddRef();
		pVUpDownAccelObj->Attach(properties.pAccelerators, acceleratorIndex);
		pVUpDownAccelObj->QueryInterface(IID_IVirtualUpDownAccelerator, reinterpret_cast<LPVOID*>(ppAccelerator));
		pVUpDownAccelObj->Release();

		return S_OK;
	}
	return DISP_E_BADINDEX;
}

STDMETHODIMP VirtualUpDownAccelerators::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP VirtualUpDownAccelerators::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.numberOfAccelerators;
	return S_OK;
}