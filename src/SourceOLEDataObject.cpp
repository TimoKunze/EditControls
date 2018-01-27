// SourceOLEDataObject.cpp: Communicates with an object through the IDataObject interface

#include "stdafx.h"
#include "SourceOLEDataObject.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP SourceOLEDataObject::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IOLEDataObject, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDataObject
STDMETHODIMP SourceOLEDataObject::DAdvise(FORMATETC* pFormatToTrack, DWORD flags, IAdviseSink* pAdviseSink, LPDWORD pConnectionID)
{
	if(!properties.pDataAdviseHolder) {
		CreateDataAdviseHolder(&properties.pDataAdviseHolder);
	}
	if(properties.pDataAdviseHolder) {
		return properties.pDataAdviseHolder->Advise(static_cast<LPDATAOBJECT>(this), pFormatToTrack, flags, pAdviseSink, pConnectionID);
	}

	return E_NOTIMPL;
}

STDMETHODIMP SourceOLEDataObject::DUnadvise(DWORD connectionID)
{
	if(properties.pDataAdviseHolder) {
		return properties.pDataAdviseHolder->Unadvise(connectionID);
	}

	return E_NOTIMPL;
}

STDMETHODIMP SourceOLEDataObject::EnumDAdvise(IEnumSTATDATA** ppEnumerator)
{
	if(properties.pDataAdviseHolder) {
		return properties.pDataAdviseHolder->EnumAdvise(ppEnumerator);
	}

	return E_NOTIMPL;
}

STDMETHODIMP SourceOLEDataObject::EnumFormatEtc(DWORD direction, IEnumFORMATETC** ppEnumerator)
{
	switch(direction) {
		case DATADIR_GET:
			Reset();
			return QueryInterface(IID_IEnumFORMATETC, reinterpret_cast<LPVOID*>(ppEnumerator));
		case DATADIR_SET:
			// TODO: Should we implement this one?
			ATLASSERT(FALSE && "SourceOLEDataObject::EnumFormatEtc() was called with DATADIR_SET");
			break;
	}

	return E_NOTIMPL;
}

STDMETHODIMP SourceOLEDataObject::GetCanonicalFormatEtc(FORMATETC* pSpecifiedFormat, FORMATETC* pGeneralFormat)
{
	if(!pSpecifiedFormat) {
		return E_INVALIDARG;
	}
	if(!pGeneralFormat) {
		return E_INVALIDARG;
	}

	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::iterator iter = properties.containedData.find(pSpecifiedFormat->cfFormat);
		if(iter == properties.containedData.end()) {
			// format not supported
			iter = properties.containedData.find(CF_TEXT);
			if(iter == properties.containedData.end()) {
				iter = properties.containedData.begin();
			}
			if(iter == properties.containedData.end()) {
				return DV_E_FORMATETC;
			}

			// we've changed cfFormat - checking dwAspect and lIndex doesn't make sense anymore
			DataEntry* pDataEntry = *(iter->second->dataEntries.begin());
			CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			return S_OK;
		}
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pSpecifiedFormat->cfFormat);
		if(!pEntry) {
			// format not supported
			pEntry = properties.containedData.Lookup(CF_TEXT);
			if(!pEntry) {
				pEntry = properties.containedData.GetAt(properties.containedData.GetStartPosition());
			}
			if(!pEntry) {
				return DV_E_FORMATETC;
			}

			// we've changed cfFormat - checking dwAspect and lIndex doesn't make sense anymore
			DataEntry* pDataEntry = pEntry->m_value->dataEntries[0];
			CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			return S_OK;
		}
	#endif

	// now look for the best dwAspect and lIndex
	BOOL init = TRUE;
	#ifdef USE_STL
		for(std::vector<DataEntry*>::iterator entryIterator = iter->second->dataEntries.begin(); entryIterator != iter->second->dataEntries.end(); ++entryIterator) {
			DataEntry* pDataEntry = *entryIterator;
	#else
		for(size_t entryIndex = 0; entryIndex < pEntry->m_value->dataEntries.GetCount(); ++entryIndex) {
			DataEntry* pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif
		if(init) {
			CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			init = FALSE;
		}

		if(pDataEntry->pFormat->dwAspect == pSpecifiedFormat->dwAspect) {
			if(pGeneralFormat->dwAspect != pSpecifiedFormat->dwAspect) {
				// we've found a format with the correct dwAspect value
				CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			}
			if(pGeneralFormat->lindex == pSpecifiedFormat->lindex) {
				// we've found an exact match - exit
				return S_OK;
			}
		}
	}

	return S_OK;
}

STDMETHODIMP SourceOLEDataObject::GetData(FORMATETC* pFormat, STGMEDIUM* pData)
{
	ATLASSERT_POINTER(pFormat, FORMATETC);
	if(!pFormat) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pData, STGMEDIUM);
	if(!pData) {
		return E_POINTER;
	}

	BOOL secondChance = FALSE;

FindDataEntry:
	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::iterator iter = properties.containedData.find(pFormat->cfFormat);
		if(iter == properties.containedData.end()) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pFormat->cfFormat);
		if(!pEntry) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#endif

	DataEntry* pDataEntry = NULL;
	BOOL found = FALSE;
	BOOL matchingTymed = FALSE;
	BOOL matchingPtd = FALSE;
	BOOL matchingAspect = FALSE;
	#ifdef USE_STL
		for(std::vector<DataEntry*>::iterator entryIterator = iter->second->dataEntries.begin(); !found && (entryIterator != iter->second->dataEntries.end()); ++entryIterator) {
			pDataEntry = *entryIterator;
	#else
		for(size_t entryIndex = 0; !found && (entryIndex < pEntry->m_value->dataEntries.GetCount()); ++entryIndex) {
			pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif

		if(pFormat->tymed & pDataEntry->pFormat->tymed) {
			matchingTymed = TRUE;
			if(pFormat->ptd == pDataEntry->pFormat->ptd) {
				matchingPtd = TRUE;
				if(pDataEntry->pFormat->dwAspect == pFormat->dwAspect) {
					matchingAspect = TRUE;
					if(pDataEntry->pFormat->lindex == pFormat->lindex) {
						found = TRUE;
					}
				}
			}
		}
	}

	if(!found) {
		// no matching entry was found
		if(matchingTymed) {
			if(matchingPtd) {
				return matchingAspect ? DV_E_LINDEX : DV_E_DVASPECT;
			} else {
				return DV_E_FORMATETC;
			}
		} else {
			return DV_E_TYMED;
		}
	}

	if(!pDataEntry->pData) {
		if(secondChance) {
			// format not supported
			return DV_E_FORMATETC;
		}

		// raise the OLESetData event
		LONG formatID = pFormat->cfFormat;
		// translate to VB's format constants
		if(static_cast<UINT>(formatID) == RegisterClipboardFormat(CF_RTF)) {
			formatID = 0xffffbf01;
		}
		if(properties.pOwnerTxtBox) {
			properties.pOwnerTxtBox->Raise_OLESetData(static_cast<IOLEDataObject*>(this), formatID, pFormat->lindex, pFormat->dwAspect);
		}

		secondChance = TRUE;
		goto FindDataEntry;
	}

	CopyStgMedium(pDataEntry->pData, pData);
	return S_OK;
}

STDMETHODIMP SourceOLEDataObject::GetDataHere(FORMATETC* /*pFormat*/, STGMEDIUM* /*pData*/)
{
	ATLASSERT(FALSE);
	return E_NOTIMPL;
}

STDMETHODIMP SourceOLEDataObject::QueryGetData(FORMATETC* pFormat)
{
	if(!pFormat) {
		return E_INVALIDARG;
	}

	BOOL secondChance = FALSE;

FindDataEntry:
	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::iterator iter = properties.containedData.find(pFormat->cfFormat);
		if(iter == properties.containedData.end()) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pFormat->cfFormat);
		if(!pEntry) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#endif

	DataEntry* pDataEntry = NULL;
	BOOL found = FALSE;
	BOOL matchingTymed = FALSE;
	BOOL matchingPtd = FALSE;
	BOOL matchingAspect = FALSE;
	#ifdef USE_STL
		for(std::vector<DataEntry*>::iterator entryIterator = iter->second->dataEntries.begin(); !found && (entryIterator != iter->second->dataEntries.end()); ++entryIterator) {
			pDataEntry = *entryIterator;
	#else
		for(size_t entryIndex = 0; !found && (entryIndex < pEntry->m_value->dataEntries.GetCount()); ++entryIndex) {
			pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif

		if(pFormat->tymed & pDataEntry->pFormat->tymed) {
			matchingTymed = TRUE;
			if(pFormat->ptd == pDataEntry->pFormat->ptd) {
				matchingPtd = TRUE;
				if(pDataEntry->pFormat->dwAspect == pFormat->dwAspect) {
					matchingAspect = TRUE;
					if(pDataEntry->pFormat->lindex == pFormat->lindex) {
						found = TRUE;
					}
				}
			}
		}
	}

	if(!found) {
		// no matching entry was found
		if(matchingTymed) {
			if(matchingPtd) {
				return matchingAspect ? DV_E_LINDEX : DV_E_DVASPECT;
			} else {
				return DV_E_FORMATETC;
			}
		} else {
			return DV_E_TYMED;
		}
	}

	if(!pDataEntry->pData) {
		if(secondChance) {
			// format not supported
			return DV_E_FORMATETC;
		}

		// raise the OLESetData event
		LONG formatID = pFormat->cfFormat;
		// translate to VB's format constants
		if(static_cast<UINT>(formatID) == RegisterClipboardFormat(CF_RTF)) {
			formatID = 0xffffbf01;
		}
		if(properties.pOwnerTxtBox) {
			properties.pOwnerTxtBox->Raise_OLESetData(static_cast<IOLEDataObject*>(this), formatID, pFormat->lindex, pFormat->dwAspect);
		}

		secondChance = TRUE;
		goto FindDataEntry;
	}

	return S_OK;
}

STDMETHODIMP SourceOLEDataObject::SetData(FORMATETC* pFormat, STGMEDIUM* pData, BOOL takeStgMediumOwnership)
{
	ATLASSERT_POINTER(pFormat, FORMATETC);
	if(!pFormat) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pData, STGMEDIUM);
	if(!pData) {
		return E_POINTER;
	}

	if(properties.numberOfItemsToDisplay > 1) {
		if(pFormat->cfFormat == static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragImageBits")))) {
			if(pData->hGlobal) {
				LPTSTR pBuffer = ConvertIntegerToString(properties.numberOfItemsToDisplay);
				if(pBuffer) {
					CT2W converter(pBuffer);
					LPWSTR pLabelText = converter;
					if(pLabelText) {
						LPVOID pDragStuff = GlobalLock(pData->hGlobal);
						SHDRAGIMAGE dragImage = {0};
						CopyMemory(&dragImage, pDragStuff, sizeof(SHDRAGIMAGE));
						LPRGBQUAD pDragImageBits = reinterpret_cast<LPRGBQUAD>(reinterpret_cast<LPBYTE>(pDragStuff) + sizeof(SHDRAGIMAGE));

						CTheme themingEngine;
						themingEngine.OpenThemeData(NULL, VSCLASS_DRAGDROP);
						if(!themingEngine.IsThemeNull()) {
							CDC memoryDC;
							memoryDC.CreateCompatibleDC();

							// calculate size and position
							DWORD textDrawStyle = DT_SINGLELINE;
							CRect textRectangle;
							themingEngine.GetThemeTextExtent(memoryDC, DD_TEXTBG, 1, pLabelText, -1, textDrawStyle | DT_CALCRECT, NULL, &textRectangle);
							MARGINS margins = {0};
							themingEngine.GetThemeMargins(memoryDC, DD_TEXTBG, 1, TMT_CONTENTMARGINS, &textRectangle, &margins);
							CRect labelRectangle = textRectangle;
							labelRectangle.left -= margins.cxLeftWidth;
							labelRectangle.right += margins.cxRightWidth;
							labelRectangle.top -= margins.cyTopHeight;
							labelRectangle.bottom += margins.cyBottomHeight;
							labelRectangle.OffsetRect(-labelRectangle.left, -labelRectangle.top);
							int xLabelStart = (dragImage.sizeDragImage.cx - labelRectangle.Width()) / 2;
							int yLabelStart = (dragImage.sizeDragImage.cy - labelRectangle.Height()) / 2;

							BITMAPINFO bitmapInfo = {0};
							bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
							bitmapInfo.bmiHeader.biWidth = labelRectangle.Width();
							bitmapInfo.bmiHeader.biHeight = labelRectangle.Height();
							bitmapInfo.bmiHeader.biPlanes = 1;
							bitmapInfo.bmiHeader.biBitCount = 32;
							bitmapInfo.bmiHeader.biCompression = BI_RGB;
							LPRGBQUAD pLabelBits = NULL;
							CBitmap dragImageBMP;
							dragImageBMP.CreateDIBSection(memoryDC, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pLabelBits), NULL, NULL);
							ATLASSERT(pLabelBits);
							if(pLabelBits) {
								// for correct alpha-blending the label bitmap must have the correct background
								LPRGBQUAD pLabelPixel = pLabelBits;
								LPRGBQUAD pDragImagePixel = pDragImageBits;
								pDragImagePixel += yLabelStart * dragImage.sizeDragImage.cx;
								for(int y = 0; y < labelRectangle.Height(); ++y, pDragImagePixel += dragImage.sizeDragImage.cx, pLabelPixel += labelRectangle.Width()) {
									CopyMemory(pLabelPixel, pDragImagePixel + xLabelStart, labelRectangle.Width() * sizeof(RGBQUAD));
								}
								HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImageBMP);

								// draw the background
								themingEngine.DrawThemeBackground(memoryDC, DD_TEXTBG, 1, &labelRectangle, NULL);

								// draw the text
								textRectangle.left = labelRectangle.left + margins.cxLeftWidth;
								textRectangle.right = labelRectangle.right - margins.cxRightWidth;
								textRectangle.top = labelRectangle.top + margins.cyTopHeight;
								textRectangle.bottom = labelRectangle.bottom - margins.cyBottomHeight;
								themingEngine.DrawThemeText(memoryDC, DD_TEXTBG, 1, pLabelText, -1, textDrawStyle, 0, &textRectangle);

								// insert the label with the drag image
								pLabelPixel = pLabelBits;
								pDragImagePixel = pDragImageBits;
								pDragImagePixel += yLabelStart * dragImage.sizeDragImage.cx;
								POINT pt = {0};
								for(pt.y = 0; pt.y < labelRectangle.Height(); ++pt.y, pDragImagePixel += dragImage.sizeDragImage.cx, pLabelPixel += labelRectangle.Width()) {
									LPRGBQUAD pDst = pDragImagePixel + xLabelStart;
									LPRGBQUAD pSrc = pLabelPixel;
									for(pt.x = 0; pt.x < labelRectangle.Width(); ++pt.x, ++pDst, ++pSrc) {
										if(pSrc->rgbReserved == 0x00) {
											// TODO: Why do we need to correct the alpha channel here?
											if(textRectangle.PtInRect(pt)) {
												pSrc->rgbReserved = 0xFF;
											}
										}
										*pDst = *pSrc;
									}
								}

								memoryDC.SelectBitmap(hPreviousBitmap);
							}
						}
						GlobalUnlock(pData->hGlobal);
					}
					SECUREFREE(pBuffer);
				}
			}
		}
	}

	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::iterator iter = properties.containedData.find(pFormat->cfFormat);
		if(iter != properties.containedData.end()) {
			std::vector<DataEntry*>::iterator iterDataEntryToReplace = iter->second->dataEntries.end();

			// Do we already store this format?
			BOOL found = FALSE;
			for(std::vector<DataEntry*>::iterator entryIterator = iter->second->dataEntries.begin(); !found && (entryIterator != iter->second->dataEntries.end()); ++entryIterator) {
				DataEntry* pDataEntry = *entryIterator;
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pFormat->cfFormat);
		if(pEntry) {
			size_t iDataEntryToReplace = static_cast<size_t>(-1);

			// Do we already store this format?
			BOOL found = FALSE;
			for(size_t entryIndex = 0; !found && (entryIndex < pEntry->m_value->dataEntries.GetCount()); ++entryIndex) {
				DataEntry* pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif

			if(pFormat->tymed & pDataEntry->pFormat->tymed) {
				if(pFormat->ptd == pDataEntry->pFormat->ptd) {
					if(pFormat->dwAspect == pDataEntry->pFormat->dwAspect) {
						if(pFormat->lindex == pDataEntry->pFormat->lindex) {
							found = TRUE;
							#ifdef USE_STL
								iterDataEntryToReplace = entryIterator;
							#else
								iDataEntryToReplace = entryIndex;
							#endif
						}
					}
				}
			}
		}

		#ifdef USE_STL
			if(iterDataEntryToReplace != iter->second->dataEntries.end()) {
				// simply replace the entry
				DataEntry* pDataEntry = *iterDataEntryToReplace;
		#else
			if(iDataEntryToReplace != -1) {
				// simply replace the entry
				DataEntry* pDataEntry = pEntry->m_value->dataEntries[iDataEntryToReplace];
		#endif
			*(pDataEntry->pFormat) = *pFormat;
			if(pDataEntry->pData) {
				ReleaseStgMedium(pDataEntry->pData);
				HeapFree(GetProcessHeap(), 0, pDataEntry->pData);
				pDataEntry->pData = NULL;
			}
			pDataEntry->pData = static_cast<LPSTGMEDIUM>(HeapAlloc(GetProcessHeap(), 0, sizeof(STGMEDIUM)));
			if(!pDataEntry->pData) {
				return E_OUTOFMEMORY;
			}
			CopyStgMedium(pData, pDataEntry->pData);
			if(takeStgMediumOwnership) {
				ReleaseStgMedium(pData);
			}

			// raise the OLESetData event
			LONG formatID = pFormat->cfFormat;
			// translate to VB's format constants
			if(static_cast<UINT>(formatID) == RegisterClipboardFormat(CF_RTF)) {
				formatID = 0xffffbf01;
			}
			if(properties.pOwnerTxtBox) {
				properties.pOwnerTxtBox->Raise_OLEReceivedNewData(static_cast<IOLEDataObject*>(this), formatID, pFormat->lindex, pFormat->dwAspect);
			}

			// done
			return S_OK;
		}
	}

	// create a new entry
	DataEntry* pDataEntry = static_cast<DataEntry*>(HeapAlloc(GetProcessHeap(), 0, sizeof(DataEntry)));
	if(pDataEntry) {
		pDataEntry->pFormat = static_cast<LPFORMATETC>(HeapAlloc(GetProcessHeap(), 0, sizeof(FORMATETC)));
		if(pDataEntry->pFormat) {
			*(pDataEntry->pFormat) = *pFormat;
			pDataEntry->pData = static_cast<LPSTGMEDIUM>(HeapAlloc(GetProcessHeap(), 0, sizeof(STGMEDIUM)));
			if(pDataEntry->pData) {
				CopyStgMedium(pData, pDataEntry->pData);
				if(takeStgMediumOwnership) {
					ReleaseStgMedium(pData);
				}
			} else {
				HeapFree(GetProcessHeap(), 0, pDataEntry->pFormat);
				HeapFree(GetProcessHeap(), 0, pDataEntry);
				return E_OUTOFMEMORY;
			}
		} else {
			HeapFree(GetProcessHeap(), 0, pDataEntry);
			return E_OUTOFMEMORY;
		}
	} else {
		return E_OUTOFMEMORY;
	}

	#ifdef USE_STL
		if(iter != properties.containedData.end()) {
			// the format already exists, so append the entry
			iter->second->dataEntries.push_back(pDataEntry);
		} else {
			// insert the format as well
			DataFormat* pDataFormat = new DataFormat;
			if(pDataFormat) {
				pDataFormat->dataEntries.push_back(pDataEntry);
				properties.containedData[pFormat->cfFormat] = pDataFormat;
			} else {
				return E_OUTOFMEMORY;
			}
		}
		ATLASSERT(properties.containedData.find(pFormat->cfFormat) != properties.containedData.end());
	#else
		if(pEntry) {
			// the format already exists, so append the entry
			pEntry->m_value->dataEntries.Add(pDataEntry);
		} else {
			// insert the format as well
			DataFormat* pDataFormat = new DataFormat;
			if(pDataFormat) {
				pDataFormat->dataEntries.Add(pDataEntry);
				properties.containedData[pFormat->cfFormat] = pDataFormat;
			} else {
				return E_OUTOFMEMORY;
			}
		}
		ATLASSERT(properties.containedData.Lookup(pFormat->cfFormat) != NULL);
	#endif

	// raise the OLESetData event
	LONG formatID = pFormat->cfFormat;
	// translate to VB's format constants
	if(static_cast<UINT>(formatID) == RegisterClipboardFormat(CF_RTF)) {
		formatID = 0xffffbf01;
	}
	if(properties.pOwnerTxtBox) {
		properties.pOwnerTxtBox->Raise_OLEReceivedNewData(static_cast<IOLEDataObject*>(this), formatID, pFormat->lindex, pFormat->dwAspect);
	}

	// done
	return S_OK;
}
// implementation of IDataObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IEnumFORMATETC
STDMETHODIMP SourceOLEDataObject::Clone(IEnumFORMATETC** /*ppEnumerator*/)
{
	return E_NOTIMPL;
}

STDMETHODIMP SourceOLEDataObject::Next(ULONG numberOfMaxFormats, FORMATETC* pFormats, ULONG* pNumberOfFormatsReturned)
{
	ATLASSERT_POINTER(pFormats, FORMATETC);
	if(!pFormats) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxFormats; ++i) {
		#ifdef USE_STL
			if(std::distance(properties.nextEnumeratedData, (std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator) properties.containedData.end()) <= 0) {
				// there's nothing more to iterate
				break;
			}
			// only enum the first data entry as approved by MSDN
			CopyMemory(&pFormats[i], (*(properties.nextEnumeratedData->second->dataEntries.begin()))->pFormat, sizeof(FORMATETC));
			++properties.nextEnumeratedData;
		#else
			if(!properties.nextEnumeratedData) {
				// there's nothing more to iterate
				break;
			}
			// only enum the first data entry as approved by MSDN
			CopyMemory(&pFormats[i], properties.containedData.GetValueAt(properties.nextEnumeratedData)->dataEntries[0]->pFormat, sizeof(FORMATETC));
			properties.containedData.GetNext(properties.nextEnumeratedData);
		#endif
	}
	if(pNumberOfFormatsReturned) {
		*pNumberOfFormatsReturned = i;
	}

	return (i == numberOfMaxFormats ? S_OK : S_FALSE);
}

STDMETHODIMP SourceOLEDataObject::Reset(void)
{
	#ifdef USE_STL
		properties.nextEnumeratedData = properties.containedData.begin();
	#else
		properties.nextEnumeratedData = properties.containedData.GetStartPosition();
	#endif
	return S_OK;
}

STDMETHODIMP SourceOLEDataObject::Skip(ULONG numberOfFormatsToSkip)
{
	#ifdef USE_STL
		std::advance(properties.nextEnumeratedData, numberOfFormatsToSkip);
	#else
		for(; (numberOfFormatsToSkip > 0) && (properties.nextEnumeratedData != NULL); --numberOfFormatsToSkip) {
			properties.containedData.GetNext(properties.nextEnumeratedData);
		}
	#endif
	return S_OK;
}
// implementation of IEnumFORMATETC
//////////////////////////////////////////////////////////////////////


SourceOLEDataObject::Properties::~Properties()
{
	#ifdef USE_STL
		for(std::unordered_map<CLIPFORMAT, DataFormat*>::iterator iter = containedData.begin(); iter != containedData.end(); ++iter) {
			ATLASSERT_POINTER(iter->second, DataFormat);
			DataFormat::Release(iter->second);
		}
		containedData.clear();
	#else
		POSITION p = containedData.GetStartPosition();
		while(p) {
			ATLASSERT_POINTER(containedData.GetValueAt(p), DataFormat);
			DataFormat::Release(containedData.GetValueAt(p));
			containedData.GetNextValue(p);
		}
		containedData.RemoveAll();
	#endif
	if(pOwnerTxtBox) {
		pOwnerTxtBox->Release();
	}
}

void SourceOLEDataObject::SetOwner(TextBox* pOwner)
{
	if(properties.pOwnerTxtBox) {
		properties.pOwnerTxtBox->Release();
	}
	properties.pOwnerTxtBox = pOwner;
	if(properties.pOwnerTxtBox) {
		properties.pOwnerTxtBox->AddRef();
	}
}


STDMETHODIMP SourceOLEDataObject::Clear(void)
{
	#ifdef USE_STL
		for(std::unordered_map<CLIPFORMAT, DataFormat*>::iterator iter = properties.containedData.begin(); iter != properties.containedData.end(); ++iter) {
			DataFormat::Release(iter->second);
		}
		properties.containedData.clear();
	#else
		POSITION p = properties.containedData.GetStartPosition();
		while(p) {
			DataFormat::Release(properties.containedData.GetValueAt(p));
			properties.containedData.GetNextValue(p);
		}
		properties.containedData.RemoveAll();
	#endif
	return S_OK;
}

STDMETHODIMP SourceOLEDataObject::GetCanonicalFormat(LONG* pFormatID, LONG* pIndex, LONG* pDataOrViewAspect)
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
	if((GetCanonicalFormatEtc(&format, pCanonicalFormat) == S_OK) && pCanonicalFormat) {
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

STDMETHODIMP SourceOLEDataObject::GetData(LONG formatID, LONG index/* = -1*/, LONG dataOrViewAspect/* = DVASPECT_CONTENT*/, VARIANT* pData/* = NULL*/)
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
	STGMEDIUM storageMedium = {0};
	HRESULT hr = GetData(&format, &storageMedium);
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
						picture.bmp.hpal = GetPaletteFromDataObject(this);
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
						picture.bmp.hpal = GetPaletteFromDataObject(this);
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

			return S_OK;

		} else if(storageMedium.tymed & TYMED_GDI) {
			if(format.cfFormat == CF_BITMAP || format.cfFormat == CF_PALETTE) {
				// the storage contains a bitmap handle - we'll copy this bitmap
				PICTDESC picture = {0};
				picture.cbSizeofstruct = sizeof(picture);
				picture.bmp.hbitmap = CopyBitmap(storageMedium.hBitmap, TRUE);
				picture.bmp.hpal = GetPaletteFromDataObject(this);
				picture.picType = PICTYPE_BITMAP;

				// now create an IDispatch object out of the bitmap
				OleCreatePictureIndirect(&picture, IID_IDispatch, TRUE, reinterpret_cast<LPVOID*>(&pData->pdispVal));
				pData->vt = VT_DISPATCH;

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

				return S_OK;
			}
		}
	}

	// format mismatch - raise VB runtime error 461
	DispatchError(hr, IID_IOLEDataObject, TEXT("OLEDataObject"), IDES_FORMATMISMATCH);
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 461);
}

STDMETHODIMP SourceOLEDataObject::GetDropDescription(VARIANT* pTargetDescription/* = NULL*/, VARIANT* pActionDescription/* = NULL*/, DropDescriptionIconConstants* pIcon/* = NULL*/)
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
	if(SUCCEEDED(GetData(&format, &medium)) && medium.hGlobal) {
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

STDMETHODIMP SourceOLEDataObject::GetFormat(LONG formatID, LONG index/* = -1*/, LONG dataOrViewAspect/* = DVASPECT_CONTENT*/, VARIANT_BOOL* pFormatAvailable/* = NULL*/)
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

	if(QueryGetData(&format) == S_OK) {
		*pFormatAvailable = VARIANT_TRUE;
	}
	return S_OK;
}

STDMETHODIMP SourceOLEDataObject::SetData(LONG formatID, VARIANT data/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, LONG index/* = -1*/, LONG dataOrViewAspect/* = DVASPECT_CONTENT*/)
{
	// translate VB's format constants
	if(formatID == 0xffffbf01) {
		formatID = RegisterClipboardFormat(CF_RTF);
	}

	FORMATETC format = {static_cast<CLIPFORMAT>(formatID), NULL, dataOrViewAspect, index, 0};
	STGMEDIUM storageMedium = {0};

	switch(format.cfFormat) {
		case CF_HDROP: {
			// the global memory must contain a pointer to an array of DROPFILES structs
			format.tymed = TYMED_HGLOBAL;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
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
			}
			break;
		}
		case CF_OEMTEXT:
		case CF_TEXT: {
			// the global memory must contain a pointer to the text
			format.tymed = TYMED_HGLOBAL;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
				if(data.vt == VT_BSTR) {
					int textSize = _bstr_t(data.bstrVal).length() * sizeof(CHAR);
					storageMedium.hGlobal = GlobalAlloc(GPTR, textSize + sizeof(CHAR));
					if(storageMedium.hGlobal) {
						LPSTR pText = static_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
						CopyMemory(pText, CW2A(data.bstrVal), textSize);
						GlobalUnlock(storageMedium.hGlobal);
					}
				}
			}
			break;
		}
		case CF_UNICODETEXT: {
			// the global memory must contain a pointer to the Unicode text
			format.tymed = TYMED_HGLOBAL;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
				if(data.vt == VT_BSTR) {
					int textSize = _bstr_t(data.bstrVal).length() * sizeof(WCHAR);
					storageMedium.hGlobal = GlobalAlloc(GPTR, textSize + sizeof(WCHAR));
					if(storageMedium.hGlobal) {
						LPWSTR pText = static_cast<LPWSTR>(GlobalLock(storageMedium.hGlobal));
						CopyMemory(pText, CW2W(data.bstrVal), textSize);
						GlobalUnlock(storageMedium.hGlobal);
					}
				}
			}
			break;
		}
		case CF_BITMAP:
		case CF_PALETTE: {
			// the storage must contain a bitmap handle - we'll copy the bitmap provided by the client
			format.tymed = TYMED_GDI;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
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
			}
			break;
		}
		case CF_ENHMETAFILE: {
			// the storage must contain an enhanced metafile handle - we'll copy the metafile provided by the client
			format.tymed = TYMED_ENHMF;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
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
			}
			break;
		}
		case CF_METAFILEPICT: {
			// the storage must contain a metafile handle - we'll copy the metafile provided by the client
			format.tymed = TYMED_MFPICT;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
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
			}
			break;
		}
		case CF_DIB: {
			// the storage must contain a BITMAPINFO struct followed by the bitmap bits
			format.tymed = TYMED_HGLOBAL;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
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
			}
			break;
		}
		case CF_DIBV5: {
			/* the storage must contain a BITMAPV5HEADER struct followed by the bitmap color space information
			   and the bitmap bits */
			format.tymed = TYMED_HGLOBAL;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
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
			}
			break;
		}
		default: {
			format.tymed = TYMED_HGLOBAL;
			if(data.vt != VT_ERROR) {
				storageMedium.tymed = format.tymed;
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
			}
			break;
		}
	}

	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::iterator iter = properties.containedData.find(format.cfFormat);
		if(iter != properties.containedData.end()) {
			std::vector<DataEntry*>::iterator iterDataEntryToReplace = iter->second->dataEntries.end();

			// Do we already store this format?
			BOOL found = FALSE;
			for(std::vector<DataEntry*>::iterator entryIterator = iter->second->dataEntries.begin(); !found && (entryIterator != iter->second->dataEntries.end()); ++entryIterator) {
				DataEntry* pDataEntry = *entryIterator;
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(format.cfFormat);
		if(pEntry) {
			size_t iDataEntryToReplace = static_cast<size_t>(-1);

			// Do we already store this format?
			BOOL found = FALSE;
			for(size_t entryIndex = 0; !found && (entryIndex < pEntry->m_value->dataEntries.GetCount()); ++entryIndex) {
				DataEntry* pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif

			if(format.tymed & pDataEntry->pFormat->tymed) {
				if(format.ptd == pDataEntry->pFormat->ptd) {
					if(format.dwAspect == pDataEntry->pFormat->dwAspect) {
						if(format.lindex == pDataEntry->pFormat->lindex) {
							found = TRUE;
							#ifdef USE_STL
								iterDataEntryToReplace = entryIterator;
							#else
								iDataEntryToReplace = entryIndex;
							#endif
						}
					}
				}
			}
		}

		#ifdef USE_STL
			if(iterDataEntryToReplace != iter->second->dataEntries.end()) {
				// simply replace the entry
				DataEntry* pDataEntry = *iterDataEntryToReplace;
		#else
			if(iDataEntryToReplace != -1) {
				// simply replace the entry
				DataEntry* pDataEntry = pEntry->m_value->dataEntries[iDataEntryToReplace];
		#endif
			*(pDataEntry->pFormat) = format;
			if(pDataEntry->pData) {
				ReleaseStgMedium(pDataEntry->pData);
				HeapFree(GetProcessHeap(), 0, pDataEntry->pData);
				pDataEntry->pData = NULL;
			}
			if(data.vt != VT_ERROR) {
				pDataEntry->pData = static_cast<LPSTGMEDIUM>(HeapAlloc(GetProcessHeap(), 0, sizeof(STGMEDIUM)));
				if(!pDataEntry->pData) {
					return E_OUTOFMEMORY;
				}
				// don't use CopyStgMedium, because we don't want to duplicate the content
				*(pDataEntry->pData) = storageMedium;
			}

			// raise the OLESetData event
			// translate to VB's format constants
			if(static_cast<UINT>(formatID) == RegisterClipboardFormat(CF_RTF)) {
				formatID = 0xffffbf01;
			}
			if(properties.pOwnerTxtBox) {
				properties.pOwnerTxtBox->Raise_OLEReceivedNewData(static_cast<IOLEDataObject*>(this), formatID, format.lindex, format.dwAspect);
			}

			// done
			return S_OK;
		}
	}

	// create a new entry
	DataEntry* pDataEntry = static_cast<DataEntry*>(HeapAlloc(GetProcessHeap(), 0, sizeof(DataEntry)));
	if(pDataEntry) {
		pDataEntry->pFormat = static_cast<LPFORMATETC>(HeapAlloc(GetProcessHeap(), 0, sizeof(FORMATETC)));
		if(pDataEntry->pFormat) {
			*(pDataEntry->pFormat) = format;
			if(data.vt != VT_ERROR) {
				pDataEntry->pData = static_cast<LPSTGMEDIUM>(HeapAlloc(GetProcessHeap(), 0, sizeof(STGMEDIUM)));
				if(pDataEntry->pData) {
					// don't use CopyStgMedium, because we don't want to duplicate the content
					*(pDataEntry->pData) = storageMedium;
				} else {
					HeapFree(GetProcessHeap(), 0, pDataEntry->pFormat);
					HeapFree(GetProcessHeap(), 0, pDataEntry);
					return E_OUTOFMEMORY;
				}
			} else {
				// HEAP_ZERO_MEMORY initializes to 0xcdcdcdcd in DEBUG mode
				pDataEntry->pData = NULL;
			}
		} else {
			HeapFree(GetProcessHeap(), 0, pDataEntry);
			return E_OUTOFMEMORY;
		}
	} else {
		return E_OUTOFMEMORY;
	}

	#ifdef USE_STL
		if(iter != properties.containedData.end()) {
			// the format already exists, so append the entry
			iter->second->dataEntries.push_back(pDataEntry);
		} else {
			// insert the format as well
			DataFormat* pDataFormat = new DataFormat;
			if(pDataFormat) {
				pDataFormat->dataEntries.push_back(pDataEntry);
				properties.containedData[format.cfFormat] = pDataFormat;
			} else {
				return E_OUTOFMEMORY;
			}
		}
	#else
		if(pEntry) {
			// the format already exists, so append the entry
			pEntry->m_value->dataEntries.Add(pDataEntry);
		} else {
			// insert the format as well
			DataFormat* pDataFormat = new DataFormat;
			if(pDataFormat) {
				pDataFormat->dataEntries.Add(pDataEntry);
				properties.containedData[format.cfFormat] = pDataFormat;
			} else {
				return E_OUTOFMEMORY;
			}
		}
	#endif

	// raise the OLESetData event
	// translate to VB's format constants
	if(static_cast<UINT>(formatID) == RegisterClipboardFormat(CF_RTF)) {
		formatID = 0xffffbf01;
	}
	if(properties.pOwnerTxtBox) {
		properties.pOwnerTxtBox->Raise_OLEReceivedNewData(static_cast<IOLEDataObject*>(this), formatID, format.lindex, format.dwAspect);
	}

	// done
	return S_OK;
}

STDMETHODIMP SourceOLEDataObject::SetDropDescription(VARIANT targetDescription/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, VARIANT actionDescription/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, DropDescriptionIconConstants icon/* = ddiNone*/)
{
	STGMEDIUM medium = {0};
	FORMATETC format = {0};

	BOOL usingThemedDragImage = FALSE;
	format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
	format.dwAspect = DVASPECT_CONTENT;
	format.lindex = -1;
	format.tymed = TYMED_HGLOBAL;
	if(SUCCEEDED(GetData(&format, &medium))) {
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
	HRESULT hr = GetData(&format, &medium);
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

	return SetData(&format, &medium, TRUE);
}