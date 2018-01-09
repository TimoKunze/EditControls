// TargetOLEDataObject.cpp: Communicates with an object through the IDataObject interface

#include "stdafx.h"
#include "TargetOLEDataObject.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TargetOLEDataObject::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IOLEDataObject, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TargetOLEDataObject::QueryIDataObjectInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(IID_IDataObject, queriedInterface)) {
		TargetOLEDataObject* pOLEDataObject = reinterpret_cast<TargetOLEDataObject*>(pThis);
		return pOLEDataObject->properties.pDataObject->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


void TargetOLEDataObject::Attach(LPDATAOBJECT pDataObject)
{
	properties.pDataObject = pDataObject;
}

void TargetOLEDataObject::Detach(void)
{
	properties.pDataObject = NULL;
}


STDMETHODIMP TargetOLEDataObject::Clear(void)
{
	// invalid object use - raise VB runtime error 425
	DispatchError(E_NOTIMPL, IID_IOLEDataObject, TEXT("OLEDataObject"), IDES_INVALIDOBJUSE);
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 425);
}

STDMETHODIMP TargetOLEDataObject::GetCanonicalFormat(LONG* pFormatID, LONG* pIndex, LONG* pDataOrViewAspect)
{
	ATLASSERT_POINTER(pFormatID, LONG);
	if(!pFormatID) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pIndex, LONG);
	if(!pIndex) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pDataOrViewAspect, LONG);
	if(!pDataOrViewAspect) {
		return E_POINTER;
	}

	// translate VB's format constants
	if(*pFormatID == 0xffffbf01) {
		*pFormatID = RegisterClipboardFormat(CF_RTF);
	}

	FORMATETC format = {static_cast<CLIPFORMAT>(*pFormatID), NULL, *pDataOrViewAspect, *pIndex, 0};
	switch(*pFormatID) {
		case CF_BITMAP:
		case CF_PALETTE:
			format.tymed = TYMED_GDI;
			break;
		case CF_ENHMETAFILE:
			format.tymed = TYMED_ENHMF;
			break;
		case CF_METAFILEPICT:
			format.tymed = TYMED_MFPICT;
			break;
		default:
			format.tymed = TYMED_HGLOBAL;
			break;
	}
	FORMATETC canonicalFormat = {0};
	LPFORMATETC pCanonicalFormat = &canonicalFormat;
	if((properties.pDataObject->GetCanonicalFormatEtc(&format, pCanonicalFormat) == S_OK) && pCanonicalFormat) {
		format = canonicalFormat;
	}
	*pFormatID = format.cfFormat;
	*pDataOrViewAspect = format.dwAspect;
	*pIndex = format.lindex;

	// translate to VB's format constants
	if(static_cast<UINT>(*pFormatID) == RegisterClipboardFormat(CF_RTF)) {
		*pFormatID = 0xffffbf01;
	}
	return S_OK;
}

STDMETHODIMP TargetOLEDataObject::GetData(LONG formatID, LONG index/* = -1*/, LONG dataOrViewAspect/* = DVASPECT_CONTENT*/, VARIANT* pData/* = NULL*/)
{
	ATLASSERT_POINTER(pData, VARIANT);
	if(!pData) {
		return E_POINTER;
	}

	VariantClear(pData);

	// translate VB's format constants
	if(formatID == 0xffffbf01) {
		formatID = RegisterClipboardFormat(CF_RTF);
	}

	if(properties.pDataObject) {
		FORMATETC format = {static_cast<CLIPFORMAT>(formatID), NULL, dataOrViewAspect, index, 0};
		switch(formatID) {
			case CF_BITMAP:
			case CF_PALETTE:
				format.tymed = TYMED_GDI;
				break;
			case CF_ENHMETAFILE:
				format.tymed = TYMED_ENHMF;
				break;
			case CF_METAFILEPICT:
				format.tymed = TYMED_MFPICT;
				break;
			default:
				format.tymed = TYMED_HGLOBAL;
				break;
		}

		// there is data available, so retrieve it
		STGMEDIUM storageMedium = {0};
		HRESULT hr = properties.pDataObject->GetData(&format, &storageMedium);
		if(hr == S_OK) {
			if(storageMedium.tymed & TYMED_HGLOBAL) {
				switch(format.cfFormat) {
					case CF_HDROP: {
						HDROP hDrop = static_cast<HDROP>(GlobalLock(storageMedium.hGlobal));
						UINT numberOfFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);

						CComSafeArray<BSTR> files;
						files.Create(numberOfFiles, 1);
						TCHAR pBuffer[MAX_PATH + 1];
						for(UINT i = 0; i < numberOfFiles; ++i) {
							DragQueryFile(hDrop, i, pBuffer, MAX_PATH);
							files.SetAt(i + 1, _bstr_t(pBuffer).Detach(), FALSE);
						}

						pData->parray = SafeArrayCreateVectorEx(VT_BSTR, 1, numberOfFiles, NULL);
						files.CopyTo(&pData->parray);
						pData->vt = VT_ARRAY | VT_BSTR;

						GlobalUnlock(storageMedium.hGlobal);
						break;
					}
					case CF_OEMTEXT:
					case CF_TEXT: {
						// the global memory contains a pointer to the text
						LPSTR pText = static_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
						pData->bstrVal = _bstr_t(pText).Detach();
						pData->vt = VT_BSTR;
						GlobalUnlock(storageMedium.hGlobal);
						break;
					}
					case CF_UNICODETEXT: {
						// the global memory contains a pointer to the Unicode text
						LPWSTR pText = static_cast<LPWSTR>(GlobalLock(storageMedium.hGlobal));
						pData->bstrVal = _bstr_t(pText).Detach();
						pData->vt = VT_BSTR;
						GlobalUnlock(storageMedium.hGlobal);
						break;
					}
					case CF_DIB: {
						// the global memory contains a BITMAPINFO struct followed by the bitmap bits
						LPBITMAPINFO pBMPInfo = static_cast<LPBITMAPINFO>(GlobalLock(storageMedium.hGlobal));
						if(pBMPInfo) {
							PICTDESC picture = {0};
							picture.cbSizeofstruct = sizeof(picture);
							picture.bmp.hbitmap = CreateBitmapFromDIB(pBMPInfo);
							picture.bmp.hpal = GetPaletteFromDataObject(properties.pDataObject);
							picture.picType = PICTYPE_BITMAP;

							// now create an IDispatch object out of the bitmap
							OleCreatePictureIndirect(&picture, IID_IDispatch, TRUE, reinterpret_cast<LPVOID*>(&pData->pdispVal));
							pData->vt = VT_DISPATCH;

							GlobalUnlock(storageMedium.hGlobal);
						}
						break;
					}
					case CF_DIBV5: {
						/* the global memory contains a BITMAPV5HEADER struct followed by the bitmap color space
							 information and the bitmap bits */
						LPBITMAPINFO pBMPInfo = static_cast<LPBITMAPINFO>(GlobalLock(storageMedium.hGlobal));
						if(pBMPInfo) {
							PICTDESC picture = {0};
							picture.cbSizeofstruct = sizeof(picture);
							picture.bmp.hbitmap = CreateBitmapFromDIB(pBMPInfo);
							picture.bmp.hpal = GetPaletteFromDataObject(properties.pDataObject);
							picture.picType = PICTYPE_BITMAP;

							// now create an IDispatch object out of the bitmap
							OleCreatePictureIndirect(&picture, IID_IDispatch, TRUE, reinterpret_cast<LPVOID*>(&pData->pdispVal));
							pData->vt = VT_DISPATCH;

							GlobalUnlock(storageMedium.hGlobal);
						}
						break;
					}
					default:
						if(format.cfFormat == RegisterClipboardFormat(CF_RTF) || format.cfFormat == RegisterClipboardFormat(TEXT("HTML Format")) || format.cfFormat == RegisterClipboardFormat(TEXT("HTML (HyperText Markup Language)"))) {
							// the global memory contains a pointer to the text
							LPSTR pText = static_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
							pData->bstrVal = _bstr_t(pText).Detach();
							pData->vt = VT_BSTR;
							GlobalUnlock(storageMedium.hGlobal);
						} else if(format.cfFormat == RegisterClipboardFormat(TEXT("text/html")) || format.cfFormat == RegisterClipboardFormat(TEXT("text/_moz_htmlcontext"))) {
							// we're smart and return it as text
							LPWSTR pHTMLText = static_cast<LPWSTR>(GlobalLock(storageMedium.hGlobal));
							pData->bstrVal = _bstr_t(pHTMLText).Detach();
							pData->vt = VT_BSTR;
							GlobalUnlock(storageMedium.hGlobal);
						} else {
							// return a byte array
							DWORD arraySize = static_cast<DWORD>(GlobalSize(storageMedium.hGlobal));
							pData->parray = SafeArrayCreateVectorEx(VT_UI1, 1, arraySize, NULL);
							LPBYTE pBinaryData = static_cast<LPBYTE>(GlobalLock(storageMedium.hGlobal));
							CopyMemory(pData->parray->pvData, pBinaryData, arraySize);
							GlobalUnlock(storageMedium.hGlobal);

							pData->vt = VT_ARRAY | VT_UI1;
						}
						break;
				}

				ReleaseStgMedium(&storageMedium);
				return S_OK;

			} else if(storageMedium.tymed & TYMED_GDI) {
				if(format.cfFormat == CF_BITMAP || format.cfFormat == CF_PALETTE) {
					// the storage contains a bitmap handle - we'll copy this bitmap
					PICTDESC picture = {0};
					picture.cbSizeofstruct = sizeof(picture);
					picture.bmp.hbitmap = CopyBitmap(storageMedium.hBitmap, TRUE);
					picture.bmp.hpal = GetPaletteFromDataObject(properties.pDataObject);
					picture.picType = PICTYPE_BITMAP;

					// now create an IDispatch object out of the bitmap
					OleCreatePictureIndirect(&picture, IID_IDispatch, TRUE, reinterpret_cast<LPVOID*>(&pData->pdispVal));
					pData->vt = VT_DISPATCH;

					ReleaseStgMedium(&storageMedium);
					return S_OK;
				}
			} else if(storageMedium.tymed & TYMED_ENHMF) {
				if(format.cfFormat == CF_ENHMETAFILE) {
					// the storage contains an enhanced metafile handle
					PICTDESC picture = {0};
					picture.cbSizeofstruct = sizeof(picture);
					picture.emf.hemf = CopyEnhMetaFile(storageMedium.hEnhMetaFile, NULL);
					picture.picType = PICTYPE_ENHMETAFILE;
					OleCreatePictureIndirect(&picture, IID_IDispatch, TRUE, reinterpret_cast<LPVOID*>(&pData->pdispVal));
					pData->vt = VT_DISPATCH;

					ReleaseStgMedium(&storageMedium);
					return S_OK;
				}
			} else if(storageMedium.tymed & TYMED_MFPICT) {
				if(format.cfFormat == CF_METAFILEPICT) {
					// the storage contains a metafile handle
					PICTDESC picture = {0};
					picture.cbSizeofstruct = sizeof(picture);
					LPMETAFILEPICT pMetaFile = static_cast<LPMETAFILEPICT>(GlobalLock(storageMedium.hMetaFilePict));
					if(pMetaFile) {
						picture.wmf.hmeta = CopyMetaFile(pMetaFile->hMF, NULL);
						picture.wmf.xExt = pMetaFile->xExt;
						picture.wmf.yExt = pMetaFile->yExt;
						GlobalUnlock(storageMedium.hMetaFilePict);
					}
					picture.picType = PICTYPE_METAFILE;
					OleCreatePictureIndirect(&picture, IID_IDispatch, TRUE, reinterpret_cast<LPVOID*>(&pData->pdispVal));
					pData->vt = VT_DISPATCH;

					ReleaseStgMedium(&storageMedium);
					return S_OK;
				}
			}
			ReleaseStgMedium(&storageMedium);
		}
		// format mismatch - raise VB runtime error 461
		DispatchError(hr, IID_IOLEDataObject, TEXT("OLEDataObject"), IDES_FORMATMISMATCH);
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 461);
	}

	return E_FAIL;
}

STDMETHODIMP TargetOLEDataObject::GetDropDescription(VARIANT* pTargetDescription/* = NULL*/, VARIANT* pActionDescription/* = NULL*/, DropDescriptionIconConstants* pIcon/* = NULL*/)
{
	if(pTargetDescription) {
		VariantInit(pTargetDescription);
	}
	if(pActionDescription) {
		VariantInit(pActionDescription);
	}
	if(pIcon) {
		*pIcon = ddiNone;
	}

	FORMATETC format = {0};
	format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_DROPDESCRIPTION));
	format.dwAspect = DVASPECT_CONTENT;
	format.lindex = -1;
	format.tymed = TYMED_HGLOBAL;
	STGMEDIUM medium = {0};
	if(SUCCEEDED(properties.pDataObject->GetData(&format, &medium)) && medium.hGlobal) {
		DROPDESCRIPTION* pDropDescription = static_cast<DROPDESCRIPTION*>(GlobalLock(medium.hGlobal));
		if(pDropDescription) {
			if(pTargetDescription) {
				pTargetDescription->vt = VT_BSTR;
				pTargetDescription->bstrVal = SysAllocString(W2OLE(pDropDescription->szInsert));
			}
			if(pActionDescription) {
				pActionDescription->vt = VT_BSTR;
				pActionDescription->bstrVal = SysAllocString(W2OLE(pDropDescription->szMessage));
			}
			if(pIcon) {
				*pIcon = static_cast<DropDescriptionIconConstants>(pDropDescription->type);
				/* HACK: The naming of the ddi* constants is a bit fucked up.
					*       On Windows Vista, DROPIMAGE_INVALID hides the whole drop description and
					*       DROPIMAGE_NOIMAGE is broken.
					*       On Windows 7, DROPIMAGE_INVALID displays the default drop description and
					*       DROPIMAGE_NOIMAGE hides the whole drop description.
					*       So the names of ddiNone and ddiUseDefault should be switched, but unfortunately we've
					*       already released controls that define ddiNone as -1.
					*/
				if(pDropDescription->type == DROPIMAGE_INVALID && RunTimeHelper::IsWin7()) {
					*pIcon = ddiUseDefault;
				} else if(pDropDescription->type == DROPIMAGE_NOIMAGE && RunTimeHelper::IsWin7()) {
					*pIcon = ddiNone;
				}
			}

			GlobalUnlock(medium.hGlobal);
		}
		ReleaseStgMedium(&medium);
	}

	return S_OK;
}

STDMETHODIMP TargetOLEDataObject::GetFormat(LONG formatID, LONG index/* = -1*/, LONG dataOrViewAspect/* = DVASPECT_CONTENT*/, VARIANT_BOOL* pFormatAvailable/* = NULL*/)
{
	ATLASSERT_POINTER(pFormatAvailable, VARIANT_BOOL);
	if(!pFormatAvailable) {
		return E_POINTER;
	}

	// translate VB's format constants
	if(formatID == 0xffffbf01) {
		formatID = RegisterClipboardFormat(CF_RTF);
	}

	*pFormatAvailable = VARIANT_FALSE;
	if(properties.pDataObject) {
		FORMATETC format = {static_cast<CLIPFORMAT>(formatID), NULL, dataOrViewAspect, index, 0};
		switch(formatID) {
			case CF_BITMAP:
			case CF_PALETTE:
				format.tymed = TYMED_GDI;
				break;
			case CF_ENHMETAFILE:
				format.tymed = TYMED_ENHMF;
				break;
			case CF_METAFILEPICT:
				format.tymed = TYMED_MFPICT;
				break;
			default:
				format.tymed = TYMED_HGLOBAL;
				break;
		}

		if(properties.pDataObject->QueryGetData(&format) == S_OK) {
			*pFormatAvailable = VARIANT_TRUE;
		}
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP TargetOLEDataObject::SetData(LONG formatID, VARIANT data/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, LONG index/* = -1*/, LONG dataOrViewAspect/* = DVASPECT_CONTENT*/)
{
	if(data.vt == VT_ERROR) {
		return E_FAIL;
	}

	// translate VB's format constants
	if(formatID == 0xffffbf01) {
		formatID = RegisterClipboardFormat(CF_RTF);
	}

	if(properties.pDataObject) {
		// transform the data into a FORMATETC/STGMEDIUM pair
		STGMEDIUM storageMedium = {0};
		FORMATETC format = {static_cast<CLIPFORMAT>(formatID), NULL, dataOrViewAspect, index, 0};
		switch(format.cfFormat) {
			case CF_HDROP: {
				// the global memory must contain a pointer to an array of DROPFILES structs
				format.tymed = storageMedium.tymed = TYMED_HGLOBAL;
				if(data.vt == VT_BSTR) {
					int textSize = _bstr_t(data.bstrVal).length() * sizeof(TCHAR);
					// entries and the list itself are null-terminated
					storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(DROPFILES) + textSize + 2 * sizeof(TCHAR));
					if(storageMedium.hGlobal) {
						LPDROPFILES pFiles = static_cast<LPDROPFILES>(GlobalLock(storageMedium.hGlobal));

						// Windows Explorer doesn't set the pt member, too
						/*::GetCursorPos(&pFiles->pt);
						ClientToScreen(&pFiles->pt);
						pFiles->fNC = TRUE;*/
						#ifdef UNICODE
							pFiles->fWide = TRUE;
						#else
							pFiles->fWide = FALSE;
						#endif
						pFiles->pFiles = sizeof(DROPFILES);

						CopyMemory(pFiles + 1, COLE2T(data.bstrVal), textSize);
						GlobalUnlock(storageMedium.hGlobal);
					}
				} else if(data.vt == (VT_ARRAY | VT_BSTR)) {
					CComSafeArray<BSTR> files;
					files.Attach(data.parray);

					// calculate the total size of memory we'll have to allocate

					// the list is null-terminated
					size_t requiredSize = sizeof(DROPFILES) + sizeof(TCHAR);
					for(LONG i = files.GetLowerBound(); i <= files.GetUpperBound(); ++i) {
						// entries are null-separated
						requiredSize += (_bstr_t(files.GetAt(i)).length() + 1) * sizeof(TCHAR);
					}

					storageMedium.hGlobal = GlobalAlloc(GPTR, requiredSize);
					if(storageMedium.hGlobal) {
						LPDROPFILES pFiles = static_cast<LPDROPFILES>(GlobalLock(storageMedium.hGlobal));

						// Windows Explorer doesn't set the pt member, too
						/*::GetCursorPos(&pFiles->pt);
						ClientToScreen(&pFiles->pt);
						pFiles->fNC = TRUE;*/
						#ifdef UNICODE
							pFiles->fWide = TRUE;
						#else
							pFiles->fWide = FALSE;
						#endif
						pFiles->pFiles = sizeof(DROPFILES);

						++pFiles;
						for(LONG i = files.GetLowerBound(); i <= files.GetUpperBound(); ++i) {
							int textSize = _bstr_t(files.GetAt(i)).length() * sizeof(TCHAR);
							CopyMemory(pFiles, COLE2T(files.GetAt(i)), textSize);
							pFiles = reinterpret_cast<LPDROPFILES>(reinterpret_cast<LPSTR>(pFiles) + textSize + sizeof(TCHAR));
						}
						GlobalUnlock(storageMedium.hGlobal);
					}

					files.Detach();
				}
				break;
			}
			case CF_OEMTEXT:
			case CF_TEXT: {
				// the global memory must contain a pointer to the text
				format.tymed = storageMedium.tymed = TYMED_HGLOBAL;
				if(data.vt == VT_BSTR) {
					int textSize = _bstr_t(data.bstrVal).length() * sizeof(CHAR);
					storageMedium.hGlobal = GlobalAlloc(GPTR, textSize + sizeof(CHAR));
					if(storageMedium.hGlobal) {
						LPSTR pText = static_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
						CopyMemory(pText, CW2A(data.bstrVal), textSize);
						GlobalUnlock(storageMedium.hGlobal);
					}
				}
				break;
			}
			case CF_UNICODETEXT: {
				// the global memory must contain a pointer to the Unicode text
				format.tymed = storageMedium.tymed = TYMED_HGLOBAL;
				if(data.vt == VT_BSTR) {
					int textSize = _bstr_t(data.bstrVal).length() * sizeof(WCHAR);
					storageMedium.hGlobal = GlobalAlloc(GPTR, textSize + sizeof(WCHAR));
					if(storageMedium.hGlobal) {
						LPWSTR pText = static_cast<LPWSTR>(GlobalLock(storageMedium.hGlobal));
						CopyMemory(pText, CW2W(data.bstrVal), textSize);
						GlobalUnlock(storageMedium.hGlobal);
					}
				}
				break;
			}
			case CF_BITMAP:
			case CF_PALETTE: {
				// the storage must contain a bitmap handle - we'll copy the bitmap provided by the client
				format.tymed = storageMedium.tymed = TYMED_GDI;
				if(data.vt == VT_DISPATCH) {
					if(data.pdispVal) {
						CComQIPtr<IPicture, &IID_IPicture> pPicture(data.pdispVal);
						if(pPicture) {
							SHORT pictureType = PICTYPE_NONE;
							pPicture->get_Type(&pictureType);
							if(pictureType == PICTYPE_BITMAP) {
								OLE_HANDLE h = NULL;
								pPicture->get_Handle(&h);

								storageMedium.hBitmap = CopyBitmap(static_cast<HBITMAP>(LongToHandle(h)));
							}
						}
					} else {
						storageMedium.hBitmap = NULL;
					}
				} else {
					VARIANT v;
					VariantInit(&v);
					if(SUCCEEDED(VariantChangeType(&v, &data, 0, VT_UI4))) {
						storageMedium.hBitmap = CopyBitmap(static_cast<HBITMAP>(LongToHandle(v.ulVal)));
					}
				}
				break;
			}
			case CF_ENHMETAFILE: {
				// the storage must contain an enhanced metafile handle - we'll copy the metafile provided by the client
				format.tymed = storageMedium.tymed = TYMED_ENHMF;
				if(data.vt == VT_DISPATCH) {
					if(data.pdispVal) {
						CComQIPtr<IPicture, &IID_IPicture> pPicture(data.pdispVal);
						if(pPicture) {
							SHORT pictureType = PICTYPE_NONE;
							pPicture->get_Type(&pictureType);
							if(pictureType == PICTYPE_ENHMETAFILE) {
								OLE_HANDLE h = NULL;
								pPicture->get_Handle(&h);

								storageMedium.hEnhMetaFile = CopyEnhMetaFile(static_cast<HENHMETAFILE>(LongToHandle(h)), NULL);
							}
						}
					} else {
						storageMedium.hEnhMetaFile = NULL;
					}
				} else {
					VARIANT v;
					VariantInit(&v);
					if(SUCCEEDED(VariantChangeType(&v, &data, 0, VT_UI4))) {
						storageMedium.hEnhMetaFile = CopyEnhMetaFile(static_cast<HENHMETAFILE>(LongToHandle(v.ulVal)), NULL);
					}
				}
				break;
			}
			case CF_METAFILEPICT: {
				// the storage must contain a metafile handle - we'll copy the metafile provided by the client
				format.tymed = storageMedium.tymed = TYMED_MFPICT;
				if(data.vt == VT_DISPATCH) {
					if(data.pdispVal) {
						CComQIPtr<IPicture, &IID_IPicture> pPicture(data.pdispVal);
						if(pPicture) {
							SHORT pictureType = PICTYPE_NONE;
							pPicture->get_Type(&pictureType);
							if(pictureType == PICTYPE_METAFILE) {
								OLE_HANDLE h = NULL;
								pPicture->get_Handle(&h);

								storageMedium.hMetaFilePict = GlobalAlloc(GPTR, sizeof(METAFILEPICT));
								if(storageMedium.hMetaFilePict) {
									LPMETAFILEPICT pMetaFile = static_cast<LPMETAFILEPICT>(GlobalLock(storageMedium.hMetaFilePict));
									pMetaFile->hMF = CopyMetaFile(static_cast<HMETAFILE>(LongToHandle(h)), NULL);
									OLE_YSIZE_HIMETRIC height = 0;
									pPicture->get_Height(&height);
									OLE_XSIZE_HIMETRIC width = 0;
									pPicture->get_Width(&width);
									pMetaFile->xExt = width;
									pMetaFile->yExt = height;
									pMetaFile->mm = MM_ANISOTROPIC;
									GlobalUnlock(storageMedium.hMetaFilePict);
								}
							}
						}
					} else {
						storageMedium.hMetaFilePict = NULL;
					}
				} else {
					VARIANT v;
					VariantInit(&v);
					if(SUCCEEDED(VariantChangeType(&v, &data, 0, VT_UI4))) {
						storageMedium.hMetaFilePict = GlobalAlloc(GPTR, sizeof(METAFILEPICT));
						if(storageMedium.hMetaFilePict) {
							LPMETAFILEPICT pMetaFile = static_cast<LPMETAFILEPICT>(GlobalLock(storageMedium.hMetaFilePict));
							pMetaFile->hMF = CopyMetaFile(static_cast<HMETAFILE>(LongToHandle(v.ulVal)), NULL);
							pMetaFile->xExt = 0;
							pMetaFile->yExt = 0;
							pMetaFile->mm = MM_ANISOTROPIC;
							GlobalUnlock(storageMedium.hMetaFilePict);
						}
					}
				}
				break;
			}
			case CF_DIB: {
				// the storage must contain a BITMAPINFO struct followed by the bitmap bits
				format.tymed = storageMedium.tymed = TYMED_HGLOBAL;
				if(data.vt == VT_DISPATCH) {
					if(data.pdispVal) {
						CComQIPtr<IPicture, &IID_IPicture> pPicture(data.pdispVal);
						if(pPicture) {
							SHORT pictureType = PICTYPE_NONE;
							pPicture->get_Type(&pictureType);
							if(pictureType == PICTYPE_BITMAP) {
								OLE_HANDLE h = NULL;
								pPicture->get_Handle(&h);
								HBITMAP hBitmap = static_cast<HBITMAP>(LongToHandle(h));
								pPicture->get_hPal(&h);

								storageMedium.hGlobal = CreateDIBFromBitmap(hBitmap, static_cast<HPALETTE>(LongToHandle(h)));
							}
						}
					} else {
						storageMedium.hGlobal = NULL;
					}
				} else {
					VARIANT v;
					VariantInit(&v);
					if(SUCCEEDED(VariantChangeType(&v, &data, 0, VT_UI4))) {
						storageMedium.hGlobal = CreateDIBFromBitmap(static_cast<HBITMAP>(LongToHandle(v.ulVal)), NULL);
					}
				}
				break;
			}
			case CF_DIBV5: {
				/* the storage must contain a BITMAPV5HEADER struct followed by the bitmap color space information
				   and the bitmap bits */
				format.tymed = storageMedium.tymed = TYMED_HGLOBAL;
				if(data.vt == VT_DISPATCH) {
					if(data.pdispVal) {
						CComQIPtr<IPicture, &IID_IPicture> pPicture(data.pdispVal);
						if(pPicture) {
							SHORT pictureType = PICTYPE_NONE;
							pPicture->get_Type(&pictureType);
							if(pictureType == PICTYPE_BITMAP) {
								OLE_HANDLE h = NULL;
								pPicture->get_Handle(&h);
								HBITMAP hBitmap = static_cast<HBITMAP>(LongToHandle(h));
								pPicture->get_hPal(&h);

								storageMedium.hGlobal = CreateDIBV5FromBitmap(hBitmap, static_cast<HPALETTE>(LongToHandle(h)));
							}
						}
					} else {
						storageMedium.hGlobal = NULL;
					}
				} else {
					VARIANT v;
					VariantInit(&v);
					if(SUCCEEDED(VariantChangeType(&v, &data, 0, VT_UI4))) {
						storageMedium.hGlobal = CreateDIBV5FromBitmap(static_cast<HBITMAP>(LongToHandle(v.ulVal)), NULL);
					}
				}
				break;
			}
			default: {
				format.tymed = storageMedium.tymed = TYMED_HGLOBAL;
				if(format.cfFormat == RegisterClipboardFormat(CF_RTF) || format.cfFormat == RegisterClipboardFormat(TEXT("HTML Format")) || format.cfFormat == RegisterClipboardFormat(TEXT("HTML (HyperText Markup Language)"))) {
					// the global memory must contain a pointer to the text
					if(data.vt == VT_BSTR) {
						int textSize = _bstr_t(data.bstrVal).length() * sizeof(CHAR);
						storageMedium.hGlobal = GlobalAlloc(GPTR, textSize + sizeof(CHAR));
						if(storageMedium.hGlobal) {
							LPSTR pText = static_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
							CopyMemory(pText, CW2A(data.bstrVal), textSize);
							GlobalUnlock(storageMedium.hGlobal);
						}
					}
				} else if(format.cfFormat == RegisterClipboardFormat(TEXT("text/html")) || format.cfFormat == RegisterClipboardFormat(TEXT("text/_moz_htmlcontext"))) {
					// the global memory must contain a pointer to the Unicode text
					if(data.vt == VT_BSTR) {
						int textSize = _bstr_t(data.bstrVal).length() * sizeof(WCHAR);
						storageMedium.hGlobal = GlobalAlloc(GPTR, textSize + sizeof(WCHAR));
						if(storageMedium.hGlobal) {
							LPWSTR pText = static_cast<LPWSTR>(GlobalLock(storageMedium.hGlobal));
							CopyMemory(pText, CW2W(data.bstrVal), textSize);
							GlobalUnlock(storageMedium.hGlobal);
						}
					}
				} else {
					// don't use VariantChangeType here as the data size might be important
					if(data.vt == (VT_ARRAY | VT_UI1)) {
						// the global memory must contain a pointer to the binary data
						DWORD arraySize = static_cast<DWORD>(data.parray->rgsabound[0].cElements * sizeof(BYTE));
						storageMedium.hGlobal = GlobalAlloc(GPTR, arraySize);
						if(storageMedium.hGlobal) {
							LPBYTE pBinaryData = static_cast<LPBYTE>(GlobalLock(storageMedium.hGlobal));
							CopyMemory(pBinaryData, data.parray->pvData, arraySize);
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_I1) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(CHAR));
						if(storageMedium.hGlobal) {
							LPSTR pData = static_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
							*pData = data.cVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_UI1) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(BYTE));
						if(storageMedium.hGlobal) {
							LPBYTE pData = static_cast<LPBYTE>(GlobalLock(storageMedium.hGlobal));
							*pData = data.bVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_I2) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(SHORT));
						if(storageMedium.hGlobal) {
							PSHORT pData = static_cast<PSHORT>(GlobalLock(storageMedium.hGlobal));
							*pData = data.iVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_UI2) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(USHORT));
						if(storageMedium.hGlobal) {
							PUSHORT pData = static_cast<PUSHORT>(GlobalLock(storageMedium.hGlobal));
							*pData = data.uiVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_I4) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(LONG));
						if(storageMedium.hGlobal) {
							PLONG pData = static_cast<PLONG>(GlobalLock(storageMedium.hGlobal));
							*pData = data.lVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_UI4) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(ULONG));
						if(storageMedium.hGlobal) {
							PULONG pData = static_cast<PULONG>(GlobalLock(storageMedium.hGlobal));
							*pData = data.ulVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_I8) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(LONGLONG));
						if(storageMedium.hGlobal) {
							PLONGLONG pData = static_cast<PLONGLONG>(GlobalLock(storageMedium.hGlobal));
							*pData = data.llVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_UI8) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(ULONGLONG));
						if(storageMedium.hGlobal) {
							PULONGLONG pData = static_cast<PULONGLONG>(GlobalLock(storageMedium.hGlobal));
							*pData = data.ullVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_INT) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(INT));
						if(storageMedium.hGlobal) {
							PINT pData = static_cast<PINT>(GlobalLock(storageMedium.hGlobal));
							*pData = data.intVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else if(data.vt == VT_UINT) {
						storageMedium.hGlobal = GlobalAlloc(GPTR, sizeof(UINT));
						if(storageMedium.hGlobal) {
							PUINT pData = static_cast<PUINT>(GlobalLock(storageMedium.hGlobal));
							*pData = data.uintVal;
							GlobalUnlock(storageMedium.hGlobal);
						}
					} else {
						// the client didn't give us valid data
						return E_FAIL;
					}
				}
				break;
			}
		}

		// finally transfer the data
		properties.pDataObject->SetData(&format, &storageMedium, TRUE);
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP TargetOLEDataObject::SetDropDescription(VARIANT targetDescription/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, VARIANT actionDescription/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, DropDescriptionIconConstants icon/* = ddiNone*/)
{
	STGMEDIUM medium = {0};
	FORMATETC format = {0};

	BOOL usingThemedDragImage = FALSE;
	format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
	format.dwAspect = DVASPECT_CONTENT;
	format.lindex = -1;
	format.tymed = TYMED_HGLOBAL;
	if(SUCCEEDED(properties.pDataObject->GetData(&format, &medium))) {
		if(medium.hGlobal) {
			LPBOOL pUsingDefaultDragImage = static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
			if(pUsingDefaultDragImage) {
				usingThemedDragImage = *pUsingDefaultDragImage;
				GlobalUnlock(medium.hGlobal);
			}
		}
		ReleaseStgMedium(&medium);
	}
	if(!usingThemedDragImage) {
		// ignore this call
		return S_OK;
	}
	ZeroMemory(&medium, sizeof(STGMEDIUM));

	DROPDESCRIPTION currentDropDescription;
	ZeroMemory(&currentDropDescription, sizeof(DROPDESCRIPTION));

	format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_DROPDESCRIPTION));
	format.dwAspect = DVASPECT_CONTENT;
	format.lindex = -1;
	format.tymed = TYMED_HGLOBAL;
	HRESULT hr = properties.pDataObject->GetData(&format, &medium);
	if(SUCCEEDED(hr)) {
		if(medium.hGlobal) {
			DROPDESCRIPTION* pDropDescription = static_cast<DROPDESCRIPTION*>(GlobalLock(medium.hGlobal));
			if(pDropDescription) {
				currentDropDescription = *pDropDescription;
			}
			GlobalUnlock(medium.hGlobal);
		}
		ReleaseStgMedium(&medium);
		ZeroMemory(&medium, sizeof(STGMEDIUM));
	}

	medium.tymed = TYMED_HGLOBAL;
	medium.hGlobal = GlobalAlloc(GPTR, sizeof(DROPDESCRIPTION));
	if(medium.hGlobal) {
		DROPDESCRIPTION* pDropDescription = static_cast<DROPDESCRIPTION*>(GlobalLock(medium.hGlobal));
		if(pDropDescription) {
			*pDropDescription = currentDropDescription;
			if(targetDescription.vt != VT_ERROR) {
				VARIANT v;
				VariantInit(&v);
				if(SUCCEEDED(VariantChangeType(&v, &targetDescription, 0, VT_BSTR))) {
					lstrcpynW(pDropDescription->szInsert, OLE2W(v.bstrVal), min(MAX_PATH, SysStringLen(v.bstrVal) + 1));
					VariantClear(&v);
				} else {
					// invalid arg - raise VB runtime error 380
					GlobalUnlock(medium.hGlobal);
					GlobalFree(medium.hGlobal);
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
			}

			if(actionDescription.vt != VT_ERROR) {
				VARIANT v;
				VariantInit(&v);
				if(SUCCEEDED(VariantChangeType(&v, &actionDescription, 0, VT_BSTR))) {
					lstrcpynW(pDropDescription->szMessage, OLE2W(v.bstrVal), min(MAX_PATH, SysStringLen(v.bstrVal) + 1));
					VariantClear(&v);
				} else {
					// invalid arg - raise VB runtime error 380
					GlobalUnlock(medium.hGlobal);
					GlobalFree(medium.hGlobal);
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
			}

			/* HACK: The naming of the ddi* constants is a bit fucked up.
			 *       On Windows Vista, DROPIMAGE_INVALID hides the whole drop description and
			 *       DROPIMAGE_NOIMAGE is broken.
			 *       On Windows 7, DROPIMAGE_INVALID displays the default drop description and
			 *       DROPIMAGE_NOIMAGE hides the whole drop description.
			 *       So the names of ddiNone and ddiUseDefault should be switched, but unfortunately we've
			 *       already released controls that define ddiNone as -1.
			 */
			if(icon == ddiUseDefault) {
				pDropDescription->type = DROPIMAGE_INVALID;
			} else if(icon == ddiNone) {
				pDropDescription->type = (RunTimeHelper::IsWin7() ? DROPIMAGE_NOIMAGE : DROPIMAGE_INVALID);
			} else {
				pDropDescription->type = static_cast<DROPIMAGETYPE>(icon);
			}

			GlobalUnlock(medium.hGlobal);
		}
	}

	return properties.pDataObject->SetData(&format, &medium, TRUE);
}