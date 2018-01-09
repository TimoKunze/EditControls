// EnumOLEVERB.cpp: An implementation of the IEnumOLEVERB interface

#include "stdafx.h"
#include "EnumOLEVERB.h"


EnumOLEVERB::EnumOLEVERB(const OLEVERB* pVerbs, int numberOfVerbs)
{
	properties.numberOfVerbs = numberOfVerbs;
	properties.pVerbs = pVerbs;

	//InterlockedIncrement(&properties.referenceCount);
}

EnumOLEVERB::~EnumOLEVERB()
{
	//InterlockedDecrement(&properties.referenceCount);
}

HRESULT EnumOLEVERB::CreateInstance(const OLEVERB* pVerbs, int numberOfVerbs, LPUNKNOWN pOuter, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(pOuter) {
		*ppInterfaceImpl = NULL;
		return CLASS_E_NOAGGREGATION;
	}

	if(numberOfVerbs > 0) {
		// create and return the IEnumOLEVERB object
		EnumOLEVERB* pEnumerator = new EnumOLEVERB(pVerbs, numberOfVerbs);
		pEnumerator->AddRef();
		HRESULT hr = pEnumerator->QueryInterface(requiredInterface, ppInterfaceImpl);
		pEnumerator->Release();
		return hr;
	} else {
		// no verbs: return NULL
		*ppInterfaceImpl = NULL;
		return OLEOBJ_E_NOVERBS;
	}
}

HRESULT EnumOLEVERB::CreateInstance(const OLEVERB* pVerbs, int numberOfVerbs, IEnumOLEVERB** ppEnumerator)
{
	return CreateInstance(pVerbs, numberOfVerbs, NULL, IID_IEnumOLEVERB, reinterpret_cast<LPVOID*>(ppEnumerator));
}


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
ULONG STDMETHODCALLTYPE EnumOLEVERB::AddRef()
{
	return InterlockedIncrement(&properties.referenceCount);
}

ULONG STDMETHODCALLTYPE EnumOLEVERB::Release()
{
	ULONG ret = InterlockedDecrement(&properties.referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		delete this;
	}
	return ret;
}

HRESULT STDMETHODCALLTYPE EnumOLEVERB::QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = this;
	} else if(requiredInterface == IID_IEnumOLEVERB) {
		*ppInterfaceImpl = this;
	} else {
		// the requested interface is not supported
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	// we return a new reference to this object, so increment the counter
	AddRef();
	return S_OK;
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IEnumOLEVERB
HRESULT STDMETHODCALLTYPE EnumOLEVERB::Next(ULONG numberOfMaxVerbs, OLEVERB* pVerbs, ULONG* pNumberOfVerbsReturned)
{
	if(!pVerbs) {
		return E_POINTER;
	}

	ULONG copiedVerbs = 0;
	for(int verb = properties.nextVerb; (verb < properties.numberOfVerbs) && (copiedVerbs < numberOfMaxVerbs); ++verb, ++copiedVerbs) {
		pVerbs[copiedVerbs] = properties.pVerbs[properties.nextVerb];
		// make a copy of the verb's name
		int bufferSize = lstrlenW(properties.pVerbs[properties.nextVerb].lpszVerbName) + 1;
		pVerbs[copiedVerbs].lpszVerbName = reinterpret_cast<LPWSTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pVerbs[copiedVerbs].lpszVerbName, bufferSize, properties.pVerbs[properties.nextVerb].lpszVerbName)));
	}
	// save the current enumeration state
	properties.nextVerb += copiedVerbs;

	// set the number of verbs that we're returning
	if(pNumberOfVerbsReturned) {
		*pNumberOfVerbsReturned = copiedVerbs;
	}

	// return S_FALSE if we return less verbs than the caller queried for
	return (copiedVerbs == numberOfMaxVerbs) ? S_OK : S_FALSE;
}

HRESULT STDMETHODCALLTYPE EnumOLEVERB::Clone(IEnumOLEVERB** ppEnumerator)
{
	HRESULT hr = CreateInstance(properties.pVerbs, properties.numberOfVerbs, NULL, IID_IEnumOLEVERB, reinterpret_cast<LPVOID*>(ppEnumerator));
	if(FAILED(hr)) {
		return hr;
	}

	// set nextVerb (NOTE: nextVerb is initialized to 0 in the constructor)
	(*ppEnumerator)->Skip(properties.nextVerb);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE EnumOLEVERB::Reset()
{
	properties.nextVerb = 0;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE EnumOLEVERB::Skip(ULONG numberOfVerbsToSkip)
{
	properties.nextVerb = min(properties.numberOfVerbs, properties.nextVerb + static_cast<int>(numberOfVerbsToSkip));
	return S_OK;
}
// implementation of IEnumOLEVERB
//////////////////////////////////////////////////////////////////////