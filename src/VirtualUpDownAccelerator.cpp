// VirtualUpDownAccelerator.cpp: A wrapper for non-existing up down accelerators.

#include "stdafx.h"
#include "VirtualUpDownAccelerator.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualUpDownAccelerator::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualUpDownAccelerator, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


void VirtualUpDownAccelerator::Attach(LPUDACCEL pAccelerators, int acceleratorIndex)
{
	properties.pAccelerators = pAccelerators;
	properties.acceleratorIndex = acceleratorIndex;
}

void VirtualUpDownAccelerator::Detach(void)
{
	properties.pAccelerators = NULL;
	properties.acceleratorIndex = -1;
}


STDMETHODIMP VirtualUpDownAccelerator::get_ActivationTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.pAccelerators[properties.acceleratorIndex].nSec;
	return S_OK;
}

STDMETHODIMP VirtualUpDownAccelerator::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.acceleratorIndex;
	return S_OK;
}

STDMETHODIMP VirtualUpDownAccelerator::get_StepSize(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.pAccelerators[properties.acceleratorIndex].nInc;
	return S_OK;
}