// TextBox.cpp: Superclasses Edit.

#include "stdafx.h"
#include "CWindowEx2.h"
#include "TextBox.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
IMEModeConstants TextBox::IMEFlags::chineseIMETable[10] = {
    imeOff,
    imeOff,
    imeOff,
    imeOff,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOff
};

IMEModeConstants TextBox::IMEFlags::japaneseIMETable[10] = {
    imeDisable,
    imeDisable,
    imeOff,
    imeOff,
    imeHiragana,
    imeHiragana,
    imeKatakana,
    imeKatakanaHalf,
    imeAlphaFull,
    imeAlpha
};

IMEModeConstants TextBox::IMEFlags::koreanIMETable[10] = {
    imeDisable,
    imeDisable,
    imeAlpha,
    imeAlpha,
    imeHangulFull,
    imeHangul,
    imeHangulFull,
    imeHangul,
    imeAlphaFull,
    imeAlpha
};

FONTDESC TextBox::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


TextBox::TextBox()
{
	properties.font.InitializePropertyWatcher(this, DISPID_TXTBOX_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_TXTBOX_MOUSEICON);
	properties.selectedTextMouseIcon.InitializePropertyWatcher(this, DISPID_TXTBOX_SELECTEDTEXTMOUSEICON);

	SIZEL size = {100, 17};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	if(RunTimeHelper::IsVista()) {
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		ATLASSUME(pWICImagingFactory);
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TextBox::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ITextBox, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP TextBox::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<TextBox>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);

		CComBSTR bstr;
		hr = pPropertyBag->Read(OLESTR("CueBanner"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_CueBanner(bstr);
		VariantClear(&propertyValue);

		bstr.Empty();
		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_Text(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP TextBox::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<TextBox>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.cueBanner.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("CueBanner"), &propertyValue);
		VariantClear(&propertyValue);

		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.text.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("Text"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP TextBox::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 27 VT_I4 properties...
	pSize->QuadPart += 27 * (sizeof(VARTYPE) + sizeof(LONG));
	// we've 1 VT_I2 property...
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(SHORT));
	// ...and 20 VT_BOOL properties...
	pSize->QuadPart += 20 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 2 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.cueBanner.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);

	// ...and 3 VT_DISPATCH properties
	pSize->QuadPart += 3 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.selectedTextMouseIcon.pPictureDisp) {
		if(FAILED(properties.selectedTextMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.selectedTextMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP TextBox::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x0C0C0C0C/*4x AppID*/) {
		// might be a legacy stream, that starts with the ATL version
		if(signature == 0x0700 || signature == 0x0710 || signature == 0x0800 || signature == 0x0900) {
			version = 0x0099;
		} else {
			return E_FAIL;
		}
	}
	//LONG version = 0;
	if(version != 0x0099) {
		if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
			return hr;
		}
		if(version > 0x0101) {
			return E_FAIL;
		}
		LONG subSignature = 0;
		if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
			return hr;
		}
		if(subSignature != 0x03030303/*4x 0x03 (-> TextBox)*/) {
			return E_FAIL;
		}
	}

	DWORD atlVersion;
	if(version == 0x0099) {
		atlVersion = 0x0900;
	} else {
		if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
			return hr;
		}
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version != 0x0100) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL acceptNumbersOnly = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL acceptTabKey = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL alwaysShowSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AutoScrollingConstants autoScrolling = static_cast<AutoScrollingConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL cancelIMECompositionOnSetFocus = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	CharacterConversionConstants characterConversion = static_cast<CharacterConversionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL completeIMECompositionOnKillFocus = propertyValue.boolVal;
	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR cueBanner;
	if(FAILED(hr = cueBanner.ReadFromStream(pStream))) {
		return hr;
	}
	if(!cueBanner) {
		cueBanner = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayCueBannerOnFocus = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL doOEMConversion = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR foreColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS formattingRectangleHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XPOS_PIXELS formattingRectangleLeft = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YPOS_PIXELS formattingRectangleTop = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS formattingRectangleWidth = propertyValue.lVal;
	if(version == 0x0100) {
		formattingRectangleHeight -= formattingRectangleTop;
		formattingRectangleWidth -= formattingRectangleLeft;
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HAlignmentConstants hAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants imeMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR insertMarkColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL insertSoftLineBreaks = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS leftMargin = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maxTextLength = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL modified = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiLine = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I2, &propertyValue))) {
		return hr;
	}
	SHORT passwordChar = propertyValue.iVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL readOnly = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS rightMargin = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ScrollBarsConstants scrollBars = static_cast<ScrollBarsConstants>(propertyValue.lVal);

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pSelectedTextMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pSelectedTextMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants selectedTextMousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS tabWidth = propertyValue.lVal;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR text;
	if(FAILED(hr = text.ReadFromStream(pStream))) {
		return hr;
	}
	if(!text) {
		text = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useCustomFormattingRectangle = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL usePasswordChar = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;

	OLE_COLOR disabledBackColor = static_cast<OLE_COLOR>(-1);
	OLE_COLOR disabledForeColor = static_cast<OLE_COLOR>(-1);
	OLEDragImageStyleConstants oleDragImageStyle = odistClassic;
	if(version >= 0x0101) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		disabledBackColor = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		disabledForeColor = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
	}


	hr = put_AcceptNumbersOnly(acceptNumbersOnly);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AcceptTabKey(acceptTabKey);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowDragDrop(allowDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AlwaysShowSelection(alwaysShowSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoScrolling(autoScrolling);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CancelIMECompositionOnSetFocus(cancelIMECompositionOnSetFocus);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CharacterConversion(characterConversion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CompleteIMECompositionOnKillFocus(completeIMECompositionOnKillFocus);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CueBanner(cueBanner);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledBackColor(disabledBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledForeColor(disabledForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayCueBannerOnFocus(displayCueBannerOnFocus);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DoOEMConversion(doOEMConversion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragScrollTimeBase(dragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleHeight(formattingRectangleHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleLeft(formattingRectangleLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleTop(formattingRectangleTop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleWidth(formattingRectangleWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HAlignment(hAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEMode(imeMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertMarkColor(insertMarkColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertSoftLineBreaks(insertSoftLineBreaks);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LeftMargin(leftMargin);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxTextLength(maxTextLength);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Modified(modified);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiLine(multiLine);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OLEDragImageStyle(oleDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_PasswordChar(passwordChar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ReadOnly(readOnly);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightMargin(rightMargin);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollBars(scrollBars);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SelectedTextMouseIcon(pSelectedTextMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SelectedTextMousePointer(selectedTextMousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TabWidth(tabWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseCustomFormattingRectangle(useCustomFormattingRectangle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UsePasswordChar(usePasswordChar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP TextBox::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x0C0C0C0C/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0101;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x03030303/*4x 0x03 (-> TextBox)*/;
	if(FAILED(hr = pStream->Write(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.acceptNumbersOnly);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.acceptTabKey);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.autoScrolling;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.cancelIMECompositionOnSetFocus);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.characterConversion;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.completeIMECompositionOnKillFocus);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	VARTYPE vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.cueBanner.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayCueBannerOnFocus);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.doOEMConversion);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.formattingRectangle.Height();
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.formattingRectangle.left;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.formattingRectangle.top;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.formattingRectangle.Width();
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.insertMarkColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.insertSoftLineBreaks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.leftMargin;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maxTextLength;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.modified);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiLine);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I2;
	propertyValue.iVal = properties.passwordChar;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.readOnly);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rightMargin;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.scrollBars;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.selectedTextMouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.selectedTextMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.selectedTextMousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.tabWidthInPixels;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.text.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useCustomFormattingRectangle);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.usePasswordChar);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0101 starts here
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND TextBox::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_STANDARD_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<TextBox>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT TextBox::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("TextBox ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void TextBox::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP TextBox::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<TextBox>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	if(properties.resetTextToName) {
		properties.resetTextToName = FALSE;

		BSTR buffer = SysAllocString(L"");
		if(SUCCEEDED(GetAmbientDisplayName(buffer))) {
			properties.text.AssignBSTR(buffer);
		} else {
			SysFreeString(buffer);
		}
	}

	return hr;
}

STDMETHODIMP TextBox::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL TextBox::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		//ATLASSERT((dialogCode & (DLGC_WANTCHARS | DLGC_HASSETSEL | DLGC_WANTARROWS)) == (DLGC_WANTCHARS | DLGC_HASSETSEL | DLGC_WANTARROWS));
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<TextBox>::PreTranslateAccelerator(pMessage, hReturnValue);
}

HIMAGELIST TextBox::CreateLegacyDragImage(int firstChar, int lastChar, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	// retrieve window details
	BOOL layoutRTL = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);

	// create the DCs we'll draw into
	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	CDC maskMemoryDC;
	maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

	// calculate the bounding rectangle of the text
	CRect textBoundingRect;
	LRESULT lr = SendMessage(EM_POSFROMCHAR, firstChar, 0);
	if(lr != -1) {
		textBoundingRect.left = LOWORD(lr);
		textBoundingRect.top = HIWORD(lr);
	}
	lr = SendMessage(EM_POSFROMCHAR, lastChar, 0);
	if(lr != -1) {
		textBoundingRect.right = LOWORD(lr);
		textBoundingRect.bottom = HIWORD(lr);
	}
	int tmp = max(textBoundingRect.left, textBoundingRect.right);
	textBoundingRect.left = min(textBoundingRect.left, textBoundingRect.right);
	textBoundingRect.right = tmp;
	tmp = max(textBoundingRect.top, textBoundingRect.bottom);
	textBoundingRect.top = min(textBoundingRect.top, textBoundingRect.bottom);
	textBoundingRect.bottom = tmp;
	if(textBoundingRect.Height() == 0) {
		LONG l = 0;
		get_LineHeight(&l);
		textBoundingRect.bottom += l;
	}

	if(pBoundingRectangle) {
		*pBoundingRectangle = textBoundingRect;
	}

	// calculate drag image size and upper-left corner
	SIZE dragImageSize = {0};
	if(pUpperLeftPoint) {
		pUpperLeftPoint->x = textBoundingRect.left;
		pUpperLeftPoint->y = textBoundingRect.top;
	}
	dragImageSize.cx = textBoundingRect.Width();
	dragImageSize.cy = textBoundingRect.Height();

	// offset RECTs
	SIZE offset = {0};
	offset.cx = textBoundingRect.left;
	offset.cy = textBoundingRect.top;
	textBoundingRect.OffsetRect(-offset.cx, -offset.cy);

	// setup the DCs we'll draw into
	memoryDC.SetBkColor(GetSysColor(COLOR_WINDOW));
	memoryDC.SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
	memoryDC.SetBkMode(TRANSPARENT);

	// create drag image bitmap
	BOOL doAlphaChannelProcessing = RunTimeHelper::IsCommCtrl6();
	BITMAPINFO bitmapInfo = {0};
	if(doAlphaChannelProcessing) {
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = dragImageSize.cx;
		bitmapInfo.bmiHeader.biHeight = -dragImageSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
	}
	CBitmap dragImage;
	LPRGBQUAD pDragImageBits = NULL;
	if(doAlphaChannelProcessing) {
		dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
	} else {
		dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageSize.cx, dragImageSize.cy);
	}
	HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
	CBitmap dragImageMask;
	dragImageMask.CreateBitmap(dragImageSize.cx, dragImageSize.cy, 1, 1, NULL);
	HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);

	// initialize the bitmap
	// we need a transparent background
	if(doAlphaChannelProcessing && pDragImageBits) {
		// we need a transparent background
		LPRGBQUAD pPixel = pDragImageBits;
		for(int y = 0; y < dragImageSize.cy; ++y) {
			for(int x = 0; x < dragImageSize.cx; ++x, ++pPixel) {
				pPixel->rgbRed = 0xFF;
				pPixel->rgbGreen = 0xFF;
				pPixel->rgbBlue = 0xFF;
				pPixel->rgbReserved = 0x00;
			}
		}
	} else {
		memoryDC.FillRect(&textBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
	}
	maskMemoryDC.FillRect(&textBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

	if(doAlphaChannelProcessing) {
		// correct the alpha channel
		LPRGBQUAD pPixel = pDragImageBits;
		POINT pt;
		for(pt.y = 0; pt.y < dragImageSize.cy; ++pt.y) {
			for(pt.x = 0; pt.x < dragImageSize.cx; ++pt.x, ++pPixel) {
				if(layoutRTL) {
					// we're working on raw data, so we've to handle WS_EX_LAYOUTRTL ourselves
					POINT pt2 = pt;
					pt2.x = dragImageSize.cx - pt.x - 1;
					if(maskMemoryDC.GetPixel(pt2.x, pt2.y) == 0x00000000) {
						// items are not themed
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				} else {
					// layout is left to right
					if(maskMemoryDC.GetPixel(pt.x, pt.y) == 0x00000000) {
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				}
			}
		}
	}

	memoryDC.SelectBitmap(hPreviousBitmap);
	maskMemoryDC.SelectBitmap(hPreviousBitmapMask);

	// create the imagelist
	HIMAGELIST hDragImageList = ImageList_Create(dragImageSize.cx, dragImageSize.cy, (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
	ImageList_SetBkColor(hDragImageList, CLR_NONE);
	ImageList_Add(hDragImageList, dragImage, dragImageMask);

	ReleaseDC(hCompatibleDC);

	return hDragImageList;
}

BOOL TextBox::CreateLegacyOLEDragImage(int firstChar, int lastChar, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(firstChar != -1 && lastChar != -1);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	POINT upperLeftPoint = {0};
	HIMAGELIST hImageList = CreateLegacyDragImage(firstChar, lastChar, &upperLeftPoint, NULL);
	if(hImageList) {
		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD flags = 0;
			pImgLst->GetItemFlags(0, &flags);
			if(flags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				pDragImage->crColorKey = 0xFFFFFFFF;

				// transfer raw data
				CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				DeleteObject(imageInfo.hbmImage);
				DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(NULL);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(NULL, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = upperLeftPoint.x + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - upperLeftPoint.x;
			}
			pDragImage->ptOffset.y = mousePosition.y - upperLeftPoint.y;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP TextBox::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			dragDropStatus.useItemCountLabelHack = FALSE;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP TextBox::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP TextBox::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP TextBox::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSource
STDMETHODIMP TextBox::GiveFeedback(DWORD effect)
{
	VARIANT_BOOL useDefaultCursors = VARIANT_TRUE;
	//if(flags.usingThemes && RunTimeHelper::IsVista()) {
		ATLASSUME(dragDropStatus.pSourceDataObject);

		BOOL isShowingLayered = FALSE;
		FORMATETC format = {0};
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("IsShowingLayered")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM medium = {0};
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				isShowingLayered = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		BOOL useDropDescriptionHack = FALSE;
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				useDropDescriptionHack = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}

		if(isShowingLayered && properties.oleDragImageStyle != odistClassic) {
			SetCursor(static_cast<HCURSOR>(LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED)));
			useDefaultCursors = VARIANT_FALSE;
		}
		if(useDropDescriptionHack) {
			// this will make drop descriptions work
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
				if(medium.hGlobal) {
					// WM_USER + 1 (with wParam = 0 and lParam = 0) hides the drag image
					#define WM_SETDROPEFFECT				WM_USER + 2     // (wParam = DCID_*, lParam = 0)
					#define DDWM_UPDATEWINDOW				WM_USER + 3     // (wParam = 0, lParam = 0)
					typedef enum DROPEFFECTS
					{
						DCID_NULL = 0,
						DCID_NO = 1,
						DCID_MOVE = 2,
						DCID_COPY = 3,
						DCID_LINK = 4,
						DCID_MAX = 5
					} DROPEFFECTS;

					HWND hWndDragWindow = *reinterpret_cast<HWND*>(GlobalLock(medium.hGlobal));
					GlobalUnlock(medium.hGlobal);

					DROPEFFECTS dropEffect = DCID_NULL;
					switch(effect) {
						case DROPEFFECT_NONE:
							dropEffect = DCID_NO;
							break;
						case DROPEFFECT_COPY:
							dropEffect = DCID_COPY;
							break;
						case DROPEFFECT_MOVE:
							dropEffect = DCID_MOVE;
							break;
						case DROPEFFECT_LINK:
							dropEffect = DCID_LINK;
							break;
					}
					if(::IsWindow(hWndDragWindow)) {
						::PostMessage(hWndDragWindow, WM_SETDROPEFFECT, dropEffect, 0);
					}
				}
				ReleaseStgMedium(&medium);
			}
		}
	//}

	Raise_OLEGiveFeedback(effect, &useDefaultCursors);
	return (useDefaultCursors == VARIANT_FALSE ? S_OK : DRAGDROP_S_USEDEFAULTCURSORS);
}

STDMETHODIMP TextBox::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
{
	HRESULT actionToContinueWith = S_OK;
	if(pressedEscape) {
		actionToContinueWith = DRAGDROP_S_CANCEL;
	} else if(!(keyState & dragDropStatus.draggingMouseButton)) {
		actionToContinueWith = DRAGDROP_S_DROP;
	}
	Raise_OLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	return actionToContinueWith;
}
// implementation of IDropSource
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSourceNotify
STDMETHODIMP TextBox::DragEnterTarget(HWND hWndTarget)
{
	Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	return S_OK;
}

STDMETHODIMP TextBox::DragLeaveTarget(void)
{
	Raise_OLEDragLeavePotentialTarget();
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP TextBox::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_TXTBOX_ALWAYSSHOWSELECTION:
		case DISPID_TXTBOX_APPEARANCE:
		case DISPID_TXTBOX_BORDERSTYLE:
		case DISPID_TXTBOX_CUEBANNER:
		case DISPID_TXTBOX_DISPLAYCUEBANNERONFOCUS:
		case DISPID_TXTBOX_FIRSTVISIBLECHAR:
		case DISPID_TXTBOX_FIRSTVISIBLELINE:
		case DISPID_TXTBOX_FORMATTINGRECTANGLEHEIGHT:
		case DISPID_TXTBOX_FORMATTINGRECTANGLELEFT:
		case DISPID_TXTBOX_FORMATTINGRECTANGLETOP:
		case DISPID_TXTBOX_FORMATTINGRECTANGLEWIDTH:
		case DISPID_TXTBOX_HALIGNMENT:
		case DISPID_TXTBOX_LASTVISIBLELINE:
		case DISPID_TXTBOX_LEFTMARGIN:
		case DISPID_TXTBOX_LINEHEIGHT:
		case DISPID_TXTBOX_MOUSEICON:
		case DISPID_TXTBOX_MOUSEPOINTER:
		case DISPID_TXTBOX_PASSWORDCHAR:
		case DISPID_TXTBOX_RIGHTMARGIN:
		case DISPID_TXTBOX_SELECTEDTEXTMOUSEICON:
		case DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER:
		case DISPID_TXTBOX_USECUSTOMFORMATTINGRECTANGLE:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_TXTBOX_ACCEPTNUMBERSONLY:
		case DISPID_TXTBOX_ACCEPTTABKEY:
		case DISPID_TXTBOX_AUTOSCROLLING:
		case DISPID_TXTBOX_CANCELIMECOMPOSITIONONSETFOCUS:
		case DISPID_TXTBOX_CHARACTERCONVERSION:
		case DISPID_TXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS:
		case DISPID_TXTBOX_DISABLEDEVENTS:
		case DISPID_TXTBOX_DONTREDRAW:
		case DISPID_TXTBOX_DOOEMCONVERSION:
		case DISPID_TXTBOX_HOVERTIME:
		case DISPID_TXTBOX_IMEMODE:
		case DISPID_TXTBOX_INSERTSOFTLINEBREAKS:
		case DISPID_TXTBOX_MAXTEXTLENGTH:
		case DISPID_TXTBOX_MULTILINE:
		case DISPID_TXTBOX_PROCESSCONTEXTMENUKEYS:
		case DISPID_TXTBOX_RIGHTTOLEFT:
		case DISPID_TXTBOX_SCROLLBARS:
		case DISPID_TXTBOX_TABSTOPS:
		case DISPID_TXTBOX_TABWIDTH:
		case DISPID_TXTBOX_USEPASSWORDCHAR:
		case DISPID_TXTBOX_WORDBREAKFUNCTION:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_TXTBOX_BACKCOLOR:
		case DISPID_TXTBOX_DISABLEDBACKCOLOR:
		case DISPID_TXTBOX_DISABLEDFORECOLOR:
		case DISPID_TXTBOX_FORECOLOR:
		case DISPID_TXTBOX_INSERTMARKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_TXTBOX_APPID:
		case DISPID_TXTBOX_APPNAME:
		case DISPID_TXTBOX_APPSHORTNAME:
		case DISPID_TXTBOX_BUILD:
		case DISPID_TXTBOX_CHARSET:
		case DISPID_TXTBOX_HDRAGIMAGELIST:
		case DISPID_TXTBOX_HWND:
		case DISPID_TXTBOX_ISRELEASE:
		case DISPID_TXTBOX_MODIFIED:
		case DISPID_TXTBOX_PROGRAMMER:
		case DISPID_TXTBOX_SELECTEDTEXT:
		case DISPID_TXTBOX_TESTER:
		case DISPID_TXTBOX_TEXT:
		case DISPID_TXTBOX_TEXTLENGTH:
		case DISPID_TXTBOX_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_TXTBOX_ALLOWDRAGDROP:
		case DISPID_TXTBOX_DRAGSCROLLTIMEBASE:
		case DISPID_TXTBOX_OLEDRAGIMAGESTYLE:
		case DISPID_TXTBOX_REGISTERFOROLEDRAGDROP:
		case DISPID_TXTBOX_SHOWDRAGIMAGE:
		case DISPID_TXTBOX_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_TXTBOX_FONT:
		case DISPID_TXTBOX_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_TXTBOX_ENABLED:
		case DISPID_TXTBOX_READONLY:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString TextBox::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString TextBox::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString TextBox::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=EditControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString TextBox::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString TextBox::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString TextBox::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP TextBox::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_TXTBOX_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_TXTBOX_AUTOSCROLLING:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.autoScrolling), description);
			break;
		case DISPID_TXTBOX_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_TXTBOX_CHARACTERCONVERSION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.characterConversion), description);
			break;
		case DISPID_TXTBOX_CUEBANNER:
		case DISPID_TXTBOX_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_TXTBOX_HALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hAlignment), description);
			break;
		case DISPID_TXTBOX_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_TXTBOX_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_TXTBOX_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_TXTBOX_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_TXTBOX_SCROLLBARS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.scrollBars), description);
			break;
		case DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.selectedTextMousePointer), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<TextBox>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP TextBox::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_TXTBOX_BORDERSTYLE:
		case DISPID_TXTBOX_OLEDRAGIMAGESTYLE:
			c = 2;
			break;
		case DISPID_TXTBOX_APPEARANCE:
		case DISPID_TXTBOX_CHARACTERCONVERSION:
		case DISPID_TXTBOX_HALIGNMENT:
			c = 3;
			break;
		case DISPID_TXTBOX_AUTOSCROLLING:
		case DISPID_TXTBOX_RIGHTTOLEFT:
		case DISPID_TXTBOX_SCROLLBARS:
			c = 4;
			break;
		case DISPID_TXTBOX_IMEMODE:
			c = 12;
			break;
		case DISPID_TXTBOX_MOUSEPOINTER:
		case DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<TextBox>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = reinterpret_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = reinterpret_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if(((property == DISPID_TXTBOX_MOUSEPOINTER) || (property == DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER)) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if(property == DISPID_TXTBOX_IMEMODE) {
			// the enum is -1-based
			--propertyValue;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = reinterpret_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_TXTBOX_APPEARANCE:
		case DISPID_TXTBOX_AUTOSCROLLING:
		case DISPID_TXTBOX_BORDERSTYLE:
		case DISPID_TXTBOX_CHARACTERCONVERSION:
		case DISPID_TXTBOX_HALIGNMENT:
		case DISPID_TXTBOX_IMEMODE:
		case DISPID_TXTBOX_MOUSEPOINTER:
		case DISPID_TXTBOX_OLEDRAGIMAGESTYLE:
		case DISPID_TXTBOX_RIGHTTOLEFT:
		case DISPID_TXTBOX_SCROLLBARS:
		case DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<TextBox>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP TextBox::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_TXTBOX_CUEBANNER:
		case DISPID_TXTBOX_TEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<TextBox>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT TextBox::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_TXTBOX_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_AUTOSCROLLING:
			switch(cookie) {
				case asNone:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGNONE, asNone);
					break;
				case asVertical:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGVERTICAL, asVertical);
					break;
				case asHorizontal:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGHORIZONTAL, asHorizontal);
					break;
				case asVertical | asHorizontal:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGBOTH, asVertical | asHorizontal);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_CHARACTERCONVERSION:
			switch(cookie) {
				case ccNone:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONNONE, ccNone);
					break;
				case ccLowerCase:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONLOWERCASE, ccLowerCase);
					break;
				case ccUpperCase:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONUPPERCASE, ccUpperCase);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_HALIGNMENT:
			switch(cookie) {
				case halLeft:
					description = GetResStringWithNumber(IDP_HALIGNMENTLEFT, halLeft);
					break;
				case halCenter:
					description = GetResStringWithNumber(IDP_HALIGNMENTCENTER, halCenter);
					break;
				case halRight:
					description = GetResStringWithNumber(IDP_HALIGNMENTRIGHT, halRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_IMEMODE:
			switch(cookie) {
				case imeInherit:
					description = GetResStringWithNumber(IDP_IMEMODEINHERIT, imeInherit);
					break;
				case imeNoControl:
					description = GetResStringWithNumber(IDP_IMEMODENOCONTROL, imeNoControl);
					break;
				case imeOn:
					description = GetResStringWithNumber(IDP_IMEMODEON, imeOn);
					break;
				case imeOff:
					description = GetResStringWithNumber(IDP_IMEMODEOFF, imeOff);
					break;
				case imeDisable:
					description = GetResStringWithNumber(IDP_IMEMODEDISABLE, imeDisable);
					break;
				case imeHiragana:
					description = GetResStringWithNumber(IDP_IMEMODEHIRAGANA, imeHiragana);
					break;
				case imeKatakana:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANA, imeKatakana);
					break;
				case imeKatakanaHalf:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANAHALF, imeKatakanaHalf);
					break;
				case imeAlphaFull:
					description = GetResStringWithNumber(IDP_IMEMODEALPHAFULL, imeAlphaFull);
					break;
				case imeAlpha:
					description = GetResStringWithNumber(IDP_IMEMODEALPHA, imeAlpha);
					break;
				case imeHangulFull:
					description = GetResStringWithNumber(IDP_IMEMODEHANGULFULL, imeHangulFull);
					break;
				case imeHangul:
					description = GetResStringWithNumber(IDP_IMEMODEHANGUL, imeHangul);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_MOUSEPOINTER:
		case DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_OLEDRAGIMAGESTYLE:
			switch(cookie) {
				case odistClassic:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLECLASSIC, odistClassic);
					break;
				case odistAeroIfAvailable:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLEAERO, odistAeroIfAvailable);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TXTBOX_SCROLLBARS:
			switch(cookie) {
				case sbNone:
					description = GetResStringWithNumber(IDP_SCROLLBARSNONE, sbNone);
					break;
				case sbVertical:
					description = GetResStringWithNumber(IDP_SCROLLBARSVERTICAL, sbVertical);
					break;
				case sbHorizontal:
					description = GetResStringWithNumber(IDP_SCROLLBARSHORIZONTAL, sbHorizontal);
					break;
				case sbVertical | sbHorizontal:
					description = GetResStringWithNumber(IDP_SCROLLBARSBOTH, sbVertical | sbHorizontal);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP TextBox::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 5;
	pPropertyPages->pElems = reinterpret_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_StockColorPage;
		pPropertyPages->pElems[3] = CLSID_StockFontPage;
		pPropertyPages->pElems[4] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP TextBox::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<TextBox>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP TextBox::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP TextBox::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT TextBox::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT TextBox::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_TXTBOX_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT TextBox::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP TextBox::get_AcceptNumbersOnly(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.acceptNumbersOnly = ((GetStyle() & ES_NUMBER) == ES_NUMBER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.acceptNumbersOnly);
	return S_OK;
}

STDMETHODIMP TextBox::put_AcceptNumbersOnly(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_ACCEPTNUMBERSONLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.acceptNumbersOnly != b) {
		properties.acceptNumbersOnly = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.acceptNumbersOnly) {
				ModifyStyle(0, ES_NUMBER);
			} else {
				ModifyStyle(ES_NUMBER, 0);
			}
		}
		FireOnChanged(DISPID_TXTBOX_ACCEPTNUMBERSONLY);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_AcceptTabKey(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.acceptTabKey);
	return S_OK;
}

STDMETHODIMP TextBox::put_AcceptTabKey(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_ACCEPTTABKEY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.acceptTabKey != b) {
		properties.acceptTabKey = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_ACCEPTTABKEY);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_AllowDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowDragDrop);
	return S_OK;
}

STDMETHODIMP TextBox::put_AllowDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_ALLOWDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowDragDrop != b) {
		properties.allowDragDrop = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_ALLOWDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_AlwaysShowSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.alwaysShowSelection = ((GetStyle() & ES_NOHIDESEL) == ES_NOHIDESEL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	return S_OK;
}

STDMETHODIMP TextBox::put_AlwaysShowSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_ALWAYSSHOWSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysShowSelection != b) {
		properties.alwaysShowSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_TXTBOX_ALWAYSSHOWSELECTION);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP TextBox::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_APPEARANCE);
	if(newValue < 0 || newValue >= aDefault) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 12;
	return S_OK;
}

STDMETHODIMP TextBox::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"EditControls");
	return S_OK;
}

STDMETHODIMP TextBox::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"EditCtls");
	return S_OK;
}

STDMETHODIMP TextBox::get_AutoScrolling(AutoScrollingConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AutoScrollingConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (ES_AUTOVSCROLL | ES_AUTOHSCROLL)) {
			case 0:
				properties.autoScrolling = asNone;
				break;
			case ES_AUTOVSCROLL:
				properties.autoScrolling = asVertical;
				break;
			case ES_AUTOHSCROLL:
				properties.autoScrolling = asHorizontal;
				break;
			case ES_AUTOVSCROLL | ES_AUTOHSCROLL:
				properties.autoScrolling = static_cast<AutoScrollingConstants>(asVertical | asHorizontal);
				break;
		}
	}

	*pValue = properties.autoScrolling;
	return S_OK;
}

STDMETHODIMP TextBox::put_AutoScrolling(AutoScrollingConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_AUTOSCROLLING);
	if(properties.autoScrolling != newValue) {
		properties.autoScrolling = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_TXTBOX_AUTOSCROLLING);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP TextBox::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD style = GetStyle();
			properties.enabled = !(style & WS_DISABLED);
			properties.readOnly = ((style & ES_READONLY) == ES_READONLY);
		}
		if(properties.enabled && !properties.readOnly) {
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP TextBox::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP TextBox::get_CancelIMECompositionOnSetFocus(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.cancelIMECompositionOnSetFocus = ((SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0) & EIMES_CANCELCOMPSTRINFOCUS) == EIMES_CANCELCOMPSTRINFOCUS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.cancelIMECompositionOnSetFocus);
	return S_OK;
}

STDMETHODIMP TextBox::put_CancelIMECompositionOnSetFocus(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_CANCELIMECOMPOSITIONONSETFOCUS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.cancelIMECompositionOnSetFocus != b) {
		properties.cancelIMECompositionOnSetFocus = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD flags = static_cast<DWORD>(SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0));
			if(properties.cancelIMECompositionOnSetFocus) {
				flags |= EIMES_CANCELCOMPSTRINFOCUS;
			} else {
				flags &= ~EIMES_CANCELCOMPSTRINFOCUS;
			}
			SendMessage(EM_SETIMESTATUS, EMSIS_COMPOSITIONSTRING, flags);
		}
		FireOnChanged(DISPID_TXTBOX_CANCELIMECOMPOSITIONONSETFOCUS);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_CharacterConversion(CharacterConversionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, CharacterConversionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & ES_LOWERCASE) {
			properties.characterConversion = ccLowerCase;
		} else if(style & ES_UPPERCASE) {
			properties.characterConversion = ccUpperCase;
		} else {
			properties.characterConversion = ccNone;
		}
	}

	*pValue = properties.characterConversion;
	return S_OK;
}

STDMETHODIMP TextBox::put_CharacterConversion(CharacterConversionConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_CHARACTERCONVERSION);
	if(properties.characterConversion != newValue) {
		properties.characterConversion = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.characterConversion) {
				case ccNone:
					ModifyStyle(ES_LOWERCASE | ES_UPPERCASE, 0);
					break;
				case ccLowerCase:
					ModifyStyle(ES_UPPERCASE, ES_LOWERCASE);
					break;
				case ccUpperCase:
					ModifyStyle(ES_LOWERCASE, ES_UPPERCASE);
					break;
			}
		}
		FireOnChanged(DISPID_TXTBOX_CHARACTERCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP TextBox::get_CompleteIMECompositionOnKillFocus(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.completeIMECompositionOnKillFocus = ((SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0) & EIMES_COMPLETECOMPSTRKILLFOCUS) == EIMES_COMPLETECOMPSTRKILLFOCUS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.completeIMECompositionOnKillFocus);
	return S_OK;
}

STDMETHODIMP TextBox::put_CompleteIMECompositionOnKillFocus(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.completeIMECompositionOnKillFocus != b) {
		properties.completeIMECompositionOnKillFocus = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD flags = static_cast<DWORD>(SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0));
			if(properties.completeIMECompositionOnKillFocus) {
				flags |= EIMES_COMPLETECOMPSTRKILLFOCUS;
			} else {
				flags &= ~EIMES_COMPLETECOMPSTRKILLFOCUS;
			}
			SendMessage(EM_SETIMESTATUS, EMSIS_COMPOSITIONSTRING, flags);
		}
		FireOnChanged(DISPID_TXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_CueBanner(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		WCHAR pBuffer[1025];
		if(SendMessage(EM_GETCUEBANNER, reinterpret_cast<WPARAM>(pBuffer), 1024)) {
			properties.cueBanner = pBuffer;
		}
	}

	*pValue = properties.cueBanner.Copy();
	return S_OK;
}

STDMETHODIMP TextBox::put_CueBanner(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_CUEBANNER);
	if(properties.cueBanner != newValue) {
		properties.cueBanner.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(EM_SETCUEBANNER, properties.displayCueBannerOnFocus, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
		}
		FireOnChanged(DISPID_TXTBOX_CUEBANNER);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_DisabledBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledBackColor;
	return S_OK;
}

STDMETHODIMP TextBox::put_DisabledBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_DISABLEDBACKCOLOR);
	if(properties.disabledBackColor != newValue) {
		properties.disabledBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD style = GetStyle();
			properties.enabled = !(style & WS_DISABLED);
			properties.readOnly = ((style & ES_READONLY) == ES_READONLY);
		}
		if(!properties.enabled || properties.readOnly) {
			if(IsWindow()) {
				CRect windowRectangle;
				GetWindowRect(&windowRectangle);
				WINDOWPOS details = {0};
				details.hwnd = *this;
				details.cx = windowRectangle.Width();
				details.cy = windowRectangle.Height();
				details.flags = SWP_DRAWFRAME | SWP_FRAMECHANGED | SWP_NOCOPYBITS | SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER;
				DefWindowProc(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&details));
			}
		}
		FireOnChanged(DISPID_TXTBOX_DISABLEDBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP TextBox::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((static_cast<long>(properties.disabledEvents) & deMouseEvents) != (static_cast<long>(newValue) & deMouseEvents)) {
			if(IsWindow()) {
				if(static_cast<long>(newValue) & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_DisabledForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledForeColor;
	return S_OK;
}

STDMETHODIMP TextBox::put_DisabledForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_DISABLEDFORECOLOR);
	if(properties.disabledForeColor != newValue) {
		properties.disabledForeColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_TXTBOX_DISABLEDFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_DisplayCueBannerOnFocus(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayCueBannerOnFocus);
	return S_OK;
}

STDMETHODIMP TextBox::put_DisplayCueBannerOnFocus(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_DISPLAYCUEBANNERONFOCUS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayCueBannerOnFocus != b) {
		properties.displayCueBannerOnFocus = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(EM_SETCUEBANNER, properties.displayCueBannerOnFocus, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
		}
		FireOnChanged(DISPID_TXTBOX_DISPLAYCUEBANNERONFOCUS);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP TextBox::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_TXTBOX_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_DoOEMConversion(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.doOEMConversion = ((GetStyle() & ES_OEMCONVERT) == ES_OEMCONVERT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.doOEMConversion);
	return S_OK;
}

STDMETHODIMP TextBox::put_DoOEMConversion(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_DOOEMCONVERSION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.doOEMConversion != b) {
		properties.doOEMConversion = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.doOEMConversion) {
				ModifyStyle(0, ES_OEMCONVERT);
			} else {
				ModifyStyle(ES_OEMCONVERT, 0);
			}
		}
		FireOnChanged(DISPID_TXTBOX_DOOEMCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_DragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP TextBox::put_DragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_DRAGSCROLLTIMEBASE);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragScrollTimeBase != newValue) {
		properties.dragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_DRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP TextBox::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_TXTBOX_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_FirstVisibleChar(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if((GetStyle() & ES_MULTILINE) == 0) {
			*pValue = static_cast<LONG>(SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0));
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_FirstVisibleLine(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(GetStyle() & ES_MULTILINE) {
			*pValue = static_cast<LONG>(SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0));
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP TextBox::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_TXTBOX_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_FONT);
	return S_OK;
}

STDMETHODIMP TextBox::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_TXTBOX_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_FONT);
	return S_OK;
}

STDMETHODIMP TextBox::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP TextBox::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_TXTBOX_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_FormattingRectangleHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.Height();
	return S_OK;
}

STDMETHODIMP TextBox::put_FormattingRectangleHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_FORMATTINGRECTANGLEHEIGHT);
	if(properties.formattingRectangle.Height() != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.bottom = properties.formattingRectangle.top + newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, 0, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_FORMATTINGRECTANGLEHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_FormattingRectangleLeft(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.left;
	return S_OK;
}

STDMETHODIMP TextBox::put_FormattingRectangleLeft(OLE_XPOS_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_FORMATTINGRECTANGLELEFT);
	if(properties.formattingRectangle.left != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.MoveToX(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, 0, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_FORMATTINGRECTANGLELEFT);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_FormattingRectangleTop(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.top;
	return S_OK;
}

STDMETHODIMP TextBox::put_FormattingRectangleTop(OLE_YPOS_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_FORMATTINGRECTANGLETOP);
	if(properties.formattingRectangle.top != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.MoveToY(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, 0, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_FORMATTINGRECTANGLETOP);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_FormattingRectangleWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.Width();
	return S_OK;
}

STDMETHODIMP TextBox::put_FormattingRectangleWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_FORMATTINGRECTANGLEWIDTH);
	if(properties.formattingRectangle.Width() != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.right = properties.formattingRectangle.left + newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, 0, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_FORMATTINGRECTANGLEWIDTH);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
			case ES_CENTER:
				properties.hAlignment = halCenter;
				break;
			case ES_RIGHT:
				properties.hAlignment = halRight;
				break;
			case ES_LEFT:
				properties.hAlignment = halLeft;
				break;
		}
	}

	*pValue = properties.hAlignment;
	return S_OK;
}

STDMETHODIMP TextBox::put_HAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_HALIGNMENT);
	if(properties.hAlignment != newValue) {
		properties.hAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(RunTimeHelper::IsVista()) {
				switch(properties.hAlignment) {
					case halLeft:
						ModifyStyle(ES_CENTER | ES_RIGHT, ES_LEFT);
						break;
					case halCenter:
						ModifyStyle(ES_LEFT | ES_RIGHT, ES_CENTER);
						break;
					case halRight:
						ModifyStyle(ES_LEFT | ES_CENTER, ES_RIGHT);
						break;
				}
				FireViewChange();
			} else {
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_TXTBOX_HALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_hDragImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(dragDropStatus.IsDragging()) {
		*pValue = HandleToLong(dragDropStatus.hDragImageList);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP TextBox::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP TextBox::get_IMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.IMEMode;
	} else {
		if((GetFocus() == *this) && (GetEffectiveIMEMode() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(*this);
			if((ime != imeInherit) && (properties.IMEMode != imeInherit)) {
				properties.IMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode();
	}
	return S_OK;
}

STDMETHODIMP TextBox::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == *this) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(*this, ime);
				}
			}
		}
		FireOnChanged(DISPID_TXTBOX_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_InsertMarkColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(EM_GETINSERTMARKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.insertMarkColor = 0;
		} else if(color != OLECOLOR2COLORREF(properties.insertMarkColor)) {
			properties.insertMarkColor = color;
		}
	}

	*pValue = properties.insertMarkColor;
	return S_OK;
}

STDMETHODIMP TextBox::put_InsertMarkColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_INSERTMARKCOLOR);
	if(properties.insertMarkColor != newValue) {
		properties.insertMarkColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_INSERTMARKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_InsertSoftLineBreaks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.insertSoftLineBreaks);
	return S_OK;
}

STDMETHODIMP TextBox::put_InsertSoftLineBreaks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_INSERTSOFTLINEBREAKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.insertSoftLineBreaks != b) {
		properties.insertSoftLineBreaks = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_FMTLINES, properties.insertSoftLineBreaks, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_INSERTSOFTLINEBREAKS);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP TextBox::get_LastVisibleLine(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(GetStyle() & ES_MULTILINE) {
			LONG lineHeight = 0;
			get_LineHeight(&lineHeight);

			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			*pValue = min(SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0) + clientRectangle.Height() / lineHeight, SendMessage(EM_GETLINECOUNT, 0, 0) - 1);
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_LeftMargin(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.leftMargin = LOWORD(SendMessage(EM_GETMARGINS, 0, 0));
	}

	*pValue = properties.leftMargin;
	return S_OK;
}

STDMETHODIMP TextBox::put_LeftMargin(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_LEFTMARGIN);
	if(properties.leftMargin != newValue) {
		properties.leftMargin = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETMARGINS, EC_LEFTMARGIN, MAKELPARAM((properties.leftMargin == -1 ? EC_USEFONTINFO : properties.leftMargin), 0));
		}
		FireOnChanged(DISPID_TXTBOX_LEFTMARGIN);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_LineHeight(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		CDCHandle targetDC = GetDC();
		HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
		HFONT hPreviousFont = targetDC.SelectFont(hFont);
		CRect rc;
		targetDC.DrawText(TEXT("A"), 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
		targetDC.SelectFont(hPreviousFont);
		ReleaseDC(targetDC);

		*pValue = rc.Height();
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_MaxTextLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.maxTextLength = static_cast<LONG>(SendMessage(EM_GETLIMITTEXT, 0, 0));
	}

	*pValue = properties.maxTextLength;
	return S_OK;
}

STDMETHODIMP TextBox::put_MaxTextLength(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_MAXTEXTLENGTH);
	if(properties.maxTextLength != newValue) {
		properties.maxTextLength = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETLIMITTEXT, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength), 0);
		}
		FireOnChanged(DISPID_TXTBOX_MAXTEXTLENGTH);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Modified(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.modified = SendMessage(EM_GETMODIFY, 0, 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.modified);
	return S_OK;
}

STDMETHODIMP TextBox::put_Modified(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_MODIFIED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.modified != b) {
		properties.modified = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETMODIFY, properties.modified, 0);
		}
		FireOnChanged(DISPID_TXTBOX_MODIFIED);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP TextBox::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TXTBOX_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_MOUSEICON);
	return S_OK;
}

STDMETHODIMP TextBox::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TXTBOX_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_MOUSEICON);
	return S_OK;
}

STDMETHODIMP TextBox::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP TextBox::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_MultiLine(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiLine = ((GetStyle() & ES_MULTILINE) == ES_MULTILINE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiLine);
	return S_OK;
}

STDMETHODIMP TextBox::put_MultiLine(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_MULTILINE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiLine != b) {
		properties.multiLine = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_TXTBOX_MULTILINE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP TextBox::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_PasswordChar(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & ES_PASSWORD) {
			properties.passwordChar = (TCHAR) SendMessage(EM_GETPASSWORDCHAR, 0, 0);
			if(static_cast<SHORT>(properties.passwordChar) == 0) {
				#ifdef UNICODE
					if(RunTimeHelper::IsCommCtrl6()) {
						properties.passwordChar = L'\x25CF';
					} else {
						properties.passwordChar = L'*';
					}
				#else
					properties.passwordChar = '*';
				#endif
			}
		}
	}

	*pValue = static_cast<SHORT>(properties.passwordChar);
	return S_OK;
}

STDMETHODIMP TextBox::put_PasswordChar(SHORT newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_PASSWORDCHAR);
	TCHAR buffer = static_cast<TCHAR>(newValue);
	if(properties.passwordChar != buffer) {
		properties.passwordChar = buffer;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.usePasswordChar) {
				if(properties.passwordChar == 0) {
					#ifdef UNICODE
						if(RunTimeHelper::IsCommCtrl6()) {
							SendMessage(EM_SETPASSWORDCHAR, L'\x25CF', 0);
						} else {
							SendMessage(EM_SETPASSWORDCHAR, L'*', 0);
						}
					#else
						SendMessage(EM_SETPASSWORDCHAR, '*', 0);
					#endif
				} else {
					SendMessage(EM_SETPASSWORDCHAR, properties.passwordChar, 0);
				}
			} else {
				SendMessage(EM_SETPASSWORDCHAR, 0, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_PASSWORDCHAR);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP TextBox::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP TextBox::get_ReadOnly(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.readOnly = ((GetStyle() & ES_READONLY) == ES_READONLY);
	}

	*pValue = BOOL2VARIANTBOOL(properties.readOnly);
	return S_OK;
}

STDMETHODIMP TextBox::put_ReadOnly(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_READONLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.readOnly != b) {
		properties.readOnly = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETREADONLY, properties.readOnly, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_READONLY);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP TextBox::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_TXTBOX_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_RightMargin(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightMargin = HIWORD(SendMessage(EM_GETMARGINS, 0, 0));
	}

	*pValue = properties.rightMargin;
	return S_OK;
}

STDMETHODIMP TextBox::put_RightMargin(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_RIGHTMARGIN);
	if(properties.rightMargin != newValue) {
		properties.rightMargin = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETMARGINS, EC_RIGHTMARGIN, MAKELPARAM(0, (properties.rightMargin == -1 ? EC_USEFONTINFO : properties.rightMargin)));
		}
		FireOnChanged(DISPID_TXTBOX_RIGHTMARGIN);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = GetExStyle();
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP TextBox::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(RunTimeHelper::IsVista()) {
				/* TODO: If we're using RTL features, we sometimes end up with an edit control that has
								 WS_EX_RIGHT set. Why? Alignment is changed, too. */
				switch(GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
					case ES_CENTER:
						properties.hAlignment = halCenter;
						break;
					case ES_RIGHT:
						properties.hAlignment = halRight;
						break;
					case ES_LEFT:
						properties.hAlignment = halLeft;
						break;
				}

				if(properties.rightToLeft & rtlLayout) {
					ModifyStyleEx(0, WS_EX_LAYOUTRTL);
				} else {
					ModifyStyleEx(WS_EX_LAYOUTRTL | WS_EX_RIGHT, 0);
				}
				if(properties.rightToLeft & rtlText) {
					ModifyStyleEx(0, WS_EX_RTLREADING);
				} else {
					ModifyStyleEx(WS_EX_RTLREADING, 0);
				}

				switch(properties.hAlignment) {
					case halLeft:
						ModifyStyle(ES_CENTER | ES_RIGHT, ES_LEFT);
						break;
					case halCenter:
						ModifyStyle(ES_LEFT | ES_RIGHT, ES_CENTER);
						break;
					case halRight:
						ModifyStyle(ES_LEFT | ES_CENTER, ES_RIGHT);
						break;
				}
				FireViewChange();
			} else {
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_TXTBOX_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_ScrollBars(ScrollBarsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ScrollBarsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (WS_VSCROLL | WS_HSCROLL)) {
			case 0:
				properties.scrollBars = sbNone;
				break;
			case WS_VSCROLL:
				properties.scrollBars = sbVertical;
				break;
			case WS_HSCROLL:
				properties.scrollBars = sbHorizontal;
				break;
			case WS_VSCROLL | WS_HSCROLL:
				properties.scrollBars = static_cast<ScrollBarsConstants>(sbVertical | sbHorizontal);
				break;
		}
	}

	*pValue = properties.scrollBars;
	return S_OK;
}

STDMETHODIMP TextBox::put_ScrollBars(ScrollBarsConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_SCROLLBARS);
	if(properties.scrollBars != newValue) {
		properties.scrollBars = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.scrollBars & sbVertical) {
				ModifyStyle(0, WS_VSCROLL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				ModifyStyle(WS_VSCROLL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			if(properties.scrollBars & sbHorizontal) {
				ModifyStyle(0, WS_HSCROLL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				ModifyStyle(WS_HSCROLL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_SCROLLBARS);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_SelectedText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		int bufferSize = GetWindowTextLength() + 1;
		LPTSTR pBuffer = new TCHAR[bufferSize];
		GetWindowText(pBuffer, bufferSize);

		int selectionStart = 0;
		int selectionEnd = 0;
		SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
		pBuffer[selectionEnd] = TEXT('\0');
		*pValue = _bstr_t(&pBuffer[selectionStart]).Detach();

		delete[] pBuffer;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::get_SelectedTextMouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.selectedTextMouseIcon.pPictureDisp) {
		properties.selectedTextMouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP TextBox::put_SelectedTextMouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TXTBOX_SELECTEDTEXTMOUSEICON);
	if(properties.selectedTextMouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.selectedTextMouseIcon.StopWatching();
		if(properties.selectedTextMouseIcon.pPictureDisp) {
			properties.selectedTextMouseIcon.pPictureDisp->Release();
			properties.selectedTextMouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.selectedTextMouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.selectedTextMouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_SELECTEDTEXTMOUSEICON);
	return S_OK;
}

STDMETHODIMP TextBox::putref_SelectedTextMouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TXTBOX_SELECTEDTEXTMOUSEICON);
	if(properties.selectedTextMouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.selectedTextMouseIcon.StopWatching();
		if(properties.selectedTextMouseIcon.pPictureDisp) {
			properties.selectedTextMouseIcon.pPictureDisp->Release();
			properties.selectedTextMouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.selectedTextMouseIcon.pPictureDisp));
		}
		properties.selectedTextMouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_SELECTEDTEXTMOUSEICON);
	return S_OK;
}

STDMETHODIMP TextBox::get_SelectedTextMousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.selectedTextMousePointer;
	return S_OK;
}

STDMETHODIMP TextBox::put_SelectedTextMousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER);
	if(properties.selectedTextMousePointer != newValue) {
		properties.selectedTextMousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	*pValue = BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP TextBox::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_SHOWDRAGIMAGE);
	if(!dragDropStatus.hDragImageList && !dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_TXTBOX_SHOWDRAGIMAGE);
	SendOnDataChange(NULL);
	return S_OK;
}

STDMETHODIMP TextBox::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP TextBox::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_TabStops(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	VariantClear(pValue);
	#ifdef USE_STL
		LONG numberOfStops = static_cast<LONG>(properties.tabStops.size());
	#else
		LONG numberOfStops = static_cast<LONG>(properties.tabStops.GetCount());
	#endif
	if(numberOfStops > 0) {
		// create the array
		pValue->vt = VT_ARRAY | VT_I4;
		pValue->parray = SafeArrayCreateVectorEx(VT_I4, 1, numberOfStops, NULL);

		HFONT hFont = NULL;
		if(IsWindow()) {
			hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
		} else {
			hFont = (properties.useSystemFont ? static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)) : properties.font.currentFont);
		}
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		HFONT hPreviousFont = memoryDC.SelectFont(hFont);
		SIZE textExtent = {0};
		memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
		int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
		memoryDC.SelectFont(hPreviousFont);
		int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

		// transfer the tab stops
		for(LONG i = 1; i <= numberOfStops; ++i) {
			// convert to pixels
			long tabStop = 2 * properties.tabStops[i - 1] * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;

			SafeArrayPutElement(pValue->parray, &i, &tabStop);
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::put_TabStops(VARIANT newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_TABSTOPS);
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		if(newValue.vt == VT_EMPTY) {
			#ifdef USE_STL
				properties.tabStops.clear();
			#else
				properties.tabStops.RemoveAll();
			#endif
			if(properties.tabWidthInPixels == -1) {
				if(SendMessage(EM_SETTABSTOPS, 0, 0)) {
					hr = S_OK;
				}
			} else {
				UINT distance = properties.tabWidthInDTUs;
				if(SendMessage(EM_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&distance))) {
					hr = S_OK;
				}
			}
			FireViewChange();
		} else if(newValue.vt & VT_ARRAY) {
			// an array
			if(newValue.parray) {
				LONG l = 0;
				SafeArrayGetLBound(newValue.parray, 1, &l);
				LONG u = 0;
				SafeArrayGetUBound(newValue.parray, 1, &u);
				if(u < l) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				UINT* pTabStops = new UINT[u - l + 1];

				HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
				CDC memoryDC;
				memoryDC.CreateCompatibleDC();
				HFONT hPreviousFont = memoryDC.SelectFont(hFont);
				SIZE textExtent = {0};
				memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
				int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
				memoryDC.SelectFont(hPreviousFont);
				int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

				VARTYPE vt = 0;
				SafeArrayGetVartype(newValue.parray, &vt);
				for(LONG i = l; i <= u; ++i) {
					UINT tabStop = 0;
					if(vt == VT_I1) {
						CHAR buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_UI1) {
						BYTE buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_I2) {
						SHORT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_UI2) {
						USHORT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_I4) {
						LONG buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_UI4) {
						ULONG buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_INT) {
						INT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_UINT) {
						UINT buffer = 0;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						tabStop = buffer;
					} else if(vt == VT_VARIANT) {
						VARIANT buffer;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						if(buffer.vt == VT_I1) {
							tabStop = buffer.cVal;
						} else if(buffer.vt == VT_UI1) {
							tabStop = buffer.bVal;
						} else if(buffer.vt == VT_I2) {
							tabStop = buffer.iVal;
						} else if(buffer.vt == VT_UI2) {
							tabStop = buffer.uiVal;
						} else if(buffer.vt == VT_I4) {
							tabStop = buffer.lVal;
						} else if(buffer.vt == VT_UI4) {
							tabStop = buffer.ulVal;
						} else if(buffer.vt == VT_INT) {
							tabStop = buffer.intVal;
						} else if(buffer.vt == VT_UINT) {
							tabStop = buffer.uintVal;
						} else {
							// invalid arg - raise VB runtime error 380
							delete[] pTabStops;
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}
					} else {
						// invalid arg - raise VB runtime error 380
						delete[] pTabStops;
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}

					// convert to dialog template units
					pTabStops[i - l] = tabStop * averageCharWidthOfSystemFont / (2 * averageCharWidthOfUsedFont);
				}
				if(SendMessage(EM_SETTABSTOPS, (u - l + 1), reinterpret_cast<LPARAM>(pTabStops))) {
					hr = S_OK;
				}
				delete[] pTabStops;
				FireViewChange();
			}
		} else {
			// a single value
			UINT tabStop = 0;
			if(newValue.vt == VT_I1) {
				tabStop = newValue.cVal;
			} else if(newValue.vt == VT_UI1) {
				tabStop = newValue.bVal;
			} else if(newValue.vt == VT_I2) {
				tabStop = newValue.iVal;
			} else if(newValue.vt == VT_UI2) {
				tabStop = newValue.uiVal;
			} else if(newValue.vt == VT_I4) {
				tabStop = newValue.lVal;
			} else if(newValue.vt == VT_UI4) {
				tabStop = newValue.ulVal;
			} else if(newValue.vt == VT_INT) {
				tabStop = newValue.intVal;
			} else if(newValue.vt == VT_UINT) {
				tabStop = newValue.uintVal;
			} else {
				// invalid arg - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}

			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			CDC memoryDC;
			memoryDC.CreateCompatibleDC();
			HFONT hPreviousFont = memoryDC.SelectFont(hFont);
			SIZE textExtent = {0};
			memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
			int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
			memoryDC.SelectFont(hPreviousFont);
			int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

			// convert to dialog template units
			tabStop = tabStop * averageCharWidthOfSystemFont / (2 * averageCharWidthOfUsedFont);

			if(SendMessage(EM_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&tabStop))) {
				hr = S_OK;
			}
			FireViewChange();
		}
	}

	FireOnChanged(DISPID_TXTBOX_TABSTOPS);
	return hr;
}

STDMETHODIMP TextBox::get_TabWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.tabWidthInPixels;
	} else {
		HFONT hFont = NULL;
		if(IsWindow()) {
			hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
		} else {
			hFont = (properties.useSystemFont ? static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)) : static_cast<HFONT>(properties.font.currentFont));
		}
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		HFONT hPreviousFont = memoryDC.SelectFont(hFont);
		SIZE textExtent = {0};
		memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
		int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
		memoryDC.SelectFont(hPreviousFont);
		int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

		// convert to pixels
		*pValue = 2 * properties.tabWidthInDTUs * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;
	}
	return S_OK;
}

STDMETHODIMP TextBox::put_TabWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_TABWIDTH);
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.tabWidthInPixels != newValue) {
		LONG oldValueInPixels = properties.tabWidthInPixels;
		properties.tabWidthInPixels = newValue;
		SetDirty(TRUE);
		LONG oldValueInDTUs = properties.tabWidthInDTUs;

		HFONT hFont = NULL;
		if(IsWindow()) {
			hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
		} else {
			hFont = (properties.useSystemFont ? static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)) : static_cast<HFONT>(properties.font.currentFont));
		}
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		HFONT hPreviousFont = memoryDC.SelectFont(hFont);
		SIZE textExtent = {0};
		memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
		int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
		memoryDC.SelectFont(hPreviousFont);
		int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

		// convert to dialog template units
		properties.tabWidthInDTUs = static_cast<LONG>(newValue) * averageCharWidthOfSystemFont / (2 * averageCharWidthOfUsedFont);
		if(IsWindow()) {
			#ifdef USE_STL
				if(properties.tabStops.empty()) {
			#else
				if(properties.tabStops.IsEmpty()) {
			#endif
				if(properties.tabWidthInPixels == -1) {
					if(!SendMessage(EM_SETTABSTOPS, 0, 0)) {
						properties.tabWidthInDTUs = oldValueInDTUs;
						properties.tabWidthInPixels = oldValueInPixels;
						return E_FAIL;
					}
				} else {
					UINT distance = properties.tabWidthInDTUs;
					if(!SendMessage(EM_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&distance))) {
						properties.tabWidthInDTUs = oldValueInDTUs;
						properties.tabWidthInPixels = oldValueInPixels;
						return E_FAIL;
					}
				}
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_TXTBOX_TABWIDTH);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP TextBox::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		GetWindowText(&properties.text);
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP TextBox::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_TEXT);
	properties.text.AssignBSTR(newValue);
	if(IsWindow()) {
		SetWindowText(COLE2CT(properties.text));
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_TEXT);
	SendOnDataChange();
	return S_OK;
}

STDMETHODIMP TextBox::get_TextLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = GetWindowTextLength();
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_UseCustomFormattingRectangle(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useCustomFormattingRectangle);
	return S_OK;
}

STDMETHODIMP TextBox::put_UseCustomFormattingRectangle(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_USECUSTOMFORMATTINGRECTANGLE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useCustomFormattingRectangle != b) {
		properties.useCustomFormattingRectangle = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, 0, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_USECUSTOMFORMATTINGRECTANGLE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_UsePasswordChar(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.usePasswordChar = ((GetStyle() & ES_PASSWORD) == ES_PASSWORD);
	}

	*pValue = BOOL2VARIANTBOOL(properties.usePasswordChar);
	return S_OK;
}

STDMETHODIMP TextBox::put_UsePasswordChar(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_USEPASSWORDCHAR);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.usePasswordChar != b) {
		properties.usePasswordChar = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.usePasswordChar) {
				if(properties.passwordChar == 0) {
					#ifdef UNICODE
						if(RunTimeHelper::IsCommCtrl6()) {
							SendMessage(EM_SETPASSWORDCHAR, L'\x25CF', 0);
						} else {
							SendMessage(EM_SETPASSWORDCHAR, L'*', 0);
						}
					#else
						SendMessage(EM_SETPASSWORDCHAR, '*', 0);
					#endif
				} else {
					SendMessage(EM_SETPASSWORDCHAR, properties.passwordChar, 0);
				}
			} else {
				SendMessage(EM_SETPASSWORDCHAR, 0, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TXTBOX_USEPASSWORDCHAR);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP TextBox::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_TXTBOX_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP TextBox::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP TextBox::get_WordBreakFunction(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(SendMessage(EM_GETWORDBREAKPROC, 0, 0));
	}
	return S_OK;
}

STDMETHODIMP TextBox::put_WordBreakFunction(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TXTBOX_WORDBREAKFUNCTION);
	if(IsWindow()) {
		SendMessage(EM_SETWORDBREAKPROC, 0, static_cast<LPARAM>(newValue));
	}

	FireOnChanged(DISPID_TXTBOX_WORDBREAKFUNCTION);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP TextBox::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP TextBox::AppendText(BSTR text, VARIANT_BOOL setCaretToEnd/* = VARIANT_FALSE*/, VARIANT_BOOL scrollToCaret/* = VARIANT_TRUE*/)
{
	if(IsWindow()) {
		BOOL invalidate = FALSE;
		if(setCaretToEnd == VARIANT_FALSE) {
			CWindowEx2(*this).InternalSetRedraw(FALSE);
			int selectionStart = 0;
			int selectionEnd = 0;
			SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
			SendMessage(EM_SETSEL, INT_MAX, INT_MAX);
			SendMessage(EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(COLE2CT(text))));
			SendMessage(EM_SETSEL, selectionStart, selectionEnd);
			CWindowEx2(*this).InternalSetRedraw(TRUE);
		} else {
			SendMessage(EM_SETSEL, INT_MAX, INT_MAX);
			SendMessage(EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(COLE2CT(text))));
			invalidate = TRUE;
		}
		if(scrollToCaret != VARIANT_FALSE) {
			SendMessage(EM_SCROLLCARET, 0, 0);
			invalidate = invalidate && TRUE;
		}
		if(invalidate && flags.useTransparentTextBackground) {
			Invalidate();
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::BeginDrag(LONG draggedTextFirstChar, LONG draggedTextLastChar, OLE_HANDLE hDragImageList/* = NULL*/, OLE_XPOS_PIXELS* pXHotSpot/* = NULL*/, OLE_YPOS_PIXELS* pYHotSpot/* = NULL*/)
{
	int xHotSpot = 0;
	if(pXHotSpot) {
		xHotSpot = *pXHotSpot;
	}
	int yHotSpot = 0;
	if(pYHotSpot) {
		yHotSpot = *pYHotSpot;
	}
	HRESULT hr = dragDropStatus.BeginDrag(this, *this, draggedTextFirstChar, draggedTextLastChar, static_cast<HIMAGELIST>(LongToHandle(hDragImageList)), &xHotSpot, &yHotSpot);
	SetCapture();
	if(pXHotSpot) {
		*pXHotSpot = xHotSpot;
	}
	if(pYHotSpot) {
		*pYHotSpot = yHotSpot;
	}

	if(dragDropStatus.hDragImageList) {
		ImageList_BeginDrag(dragDropStatus.hDragImageList, 0, xHotSpot, yHotSpot);
		dragDropStatus.dragImageIsHidden = 0;
		ImageList_DragEnter(0, 0, 0);

		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}
	return hr;
}

STDMETHODIMP TextBox::CanUndo(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = BOOL2VARIANTBOOL(SendMessage(EM_CANUNDO, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::CharIndexToPosition(LONG characterIndex, OLE_XPOS_PIXELS* pX/* = NULL*/, OLE_YPOS_PIXELS* pY/* = NULL*/)
{
	if(IsWindow()) {
		LRESULT lr = SendMessage(EM_POSFROMCHAR, characterIndex, 0);
		if(lr != -1) {
			if(pX) {
				*pX = LOWORD(lr);
			}
			if(pY) {
				*pY = HIWORD(lr);
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::EmptyUndoBuffer(void)
{
	if(IsWindow()) {
		SendMessage(EM_EMPTYUNDOBUFFER, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::EndDrag(VARIANT_BOOL abort)
{
	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	KillTimer(timers.ID_DRAGSCROLL);
	ReleaseCapture();
	if(dragDropStatus.hDragImageList) {
		dragDropStatus.HideDragImage(TRUE);
		ImageList_EndDrag();
	}

	HRESULT hr = S_OK;
	if(abort) {
		hr = Raise_AbortedDrag();
	} else {
		hr = Raise_Drop();
	}

	dragDropStatus.EndDrag();
	Invalidate();

	return hr;
}

STDMETHODIMP TextBox::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP TextBox::GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, LONG* pCharacterIndex)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pCharacterIndex, LONG);
	if(!pCharacterIndex) {
		return E_POINTER;
	}

	*pCharacterIndex = -1;
	*pRelativePosition = impNowhere;

	LRESULT lr = SendMessage(EM_CHARFROMPOS, 0, MAKELPARAM(x, y));
	if((LOWORD(lr) != 65535) || (HIWORD(lr) != 65535)) {
		*pCharacterIndex = LOWORD(lr);
		*pRelativePosition = impBefore;

		int textLength = GetWindowTextLength();
		if(*pCharacterIndex == textLength) {
			/* Under some conditions we end up suggesting to put the insertion mark after a line break. The
			   result is too much white space between the character and the insertion mark. So we detect
			   such cases and adjust the proposed position. */
			LPTSTR pBuffer = new TCHAR[textLength + 1];
			GetWindowText(pBuffer, textLength + 1);
			do {
				if(*pCharacterIndex > 0) {
					--(*pCharacterIndex);
					*pRelativePosition = impAfter;
				} else {
					break;
				}
			} while((pBuffer[*pCharacterIndex] == 0x0A) || (pBuffer[*pCharacterIndex] == 0x0D));
			delete[] pBuffer;
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::GetDraggedTextRange(LONG* pDraggedTextFirstChar/* = NULL*/, LONG* pDraggedTextLastChar/* = NULL*/)
{
	if(pDraggedTextFirstChar) {
		*pDraggedTextFirstChar = dragDropStatus.draggedTextFirstChar;
	}
	if(pDraggedTextLastChar) {
		*pDraggedTextLastChar = dragDropStatus.draggedTextLastChar;
	}
	return S_OK;
}

STDMETHODIMP TextBox::GetFirstCharOfLine(LONG lineIndex, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(SendMessage(EM_LINEINDEX, lineIndex, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, LONG* pCharacterIndex)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pCharacterIndex, LONG);
	if(!pCharacterIndex) {
		return E_POINTER;
	}

	*pCharacterIndex = insertMark.characterIndex;
	if(insertMark.characterIndex != -1) {
		if(insertMark.afterChar) {
			*pRelativePosition = impAfter;
		} else {
			*pRelativePosition = impBefore;
		}
	} else {
		*pRelativePosition = impNowhere;
	}

	return S_OK;
}

STDMETHODIMP TextBox::GetLine(LONG lineIndex, BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		// HACK: EM_LINELENGTH doesn't work as expected, e. g. it may return garbage if ES_CENTER is set
		//WORD bufferSize = static_cast<WORD>(SendMessage(EM_LINELENGTH, lineIndex, 0));
		WORD bufferSize = 2048;
		LPTSTR pBuffer = new TCHAR[bufferSize + 1];
		*reinterpret_cast<PWORD>(pBuffer) = bufferSize;
		int copied = static_cast<int>(SendMessage(EM_GETLINE, lineIndex, reinterpret_cast<LPARAM>(pBuffer)));

		if(copied > 0) {
			pBuffer[copied] = TEXT('\0');

			CT2OLE converter(pBuffer);
			*pValue = SysAllocStringLen(converter, copied);
		} else {
			*pValue = SysAllocStringLen(NULL, 0);
		}
		delete[] pBuffer;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::GetLineCount(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(SendMessage(EM_GETLINECOUNT, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::GetLineFromChar(LONG characterIndex, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(SendMessage(EM_LINEFROMCHAR, characterIndex, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::GetLineLength(LONG lineIndex, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		// HACK: EM_LINELENGTH doesn't work as expected, e. g. it may return garbage if ES_CENTER is set
		//*pValue = static_cast<LONG>(SendMessage(EM_LINELENGTH, lineIndex, 0));
		WORD bufferSize = 2048;
		LPTSTR pBuffer = new TCHAR[bufferSize + 1];
		*reinterpret_cast<PWORD>(pBuffer) = bufferSize;
		*pValue = static_cast<LONG>(SendMessage(EM_GETLINE, lineIndex, reinterpret_cast<LPARAM>(pBuffer)));
		delete[] pBuffer;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::GetSelection(LONG* pSelectionStart/* = NULL*/, LONG* pSelectionEnd/* = NULL*/)
{
	if(IsWindow()) {
		SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(pSelectionStart), reinterpret_cast<LPARAM>(pSelectionEnd));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::HideBalloonTip(VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(EM_HIDEBALLOONTIP, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::IsTextTruncated(LONG lineIndex/* = 0*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = VARIANT_FALSE;

		// HACK: EM_LINELENGTH doesn't work as expected, e. g. it may return garbage if ES_CENTER is set
		//bufferSize = static_cast<LONG>(SendMessage(EM_LINELENGTH, lineIndex, 0));
		WORD bufferSize = 2048;
		LPTSTR pBuffer = new TCHAR[bufferSize + 1];
		*reinterpret_cast<PWORD>(pBuffer) = bufferSize;
		int copied = static_cast<int>(SendMessage(EM_GETLINE, lineIndex, reinterpret_cast<LPARAM>(pBuffer)));

		if(copied > 0) {
			pBuffer[copied] = TEXT('\0');

			// calculate the width
			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			DWORD margins = static_cast<DWORD>(SendMessage(EM_GETMARGINS, 0, 0));
			clientRectangle.left += LOWORD(margins);
			clientRectangle.right -= HIWORD(margins);

			int firstCharacterInLine = static_cast<int>(SendMessage(EM_LINEINDEX, lineIndex, 0));
			int lastCharacterInLine = firstCharacterInLine + copied;
			CRect textRectangle;
			textRectangle.left = GET_X_LPARAM(SendMessage(EM_POSFROMCHAR, firstCharacterInLine, 0));
			textRectangle.right = GET_X_LPARAM(SendMessage(EM_POSFROMCHAR, lastCharacterInLine, 0));
			if(textRectangle.right < textRectangle.left) {
				// it's the last line and it doesn't end with a line break
				--lastCharacterInLine;
				textRectangle.right = GET_X_LPARAM(SendMessage(EM_POSFROMCHAR, lastCharacterInLine, 0));

				// add the last character's width
				CDCHandle targetDC = GetDC();
				HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
				HFONT hPreviousFont = targetDC.SelectFont(hFont);

				CRect rc = clientRectangle;
				switch(GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
					case ES_CENTER:
						targetDC.DrawText(&pBuffer[copied - 1], 1, &rc, DT_CALCRECT | DT_CENTER | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
						break;
					case ES_RIGHT:
						targetDC.DrawText(&pBuffer[copied - 1], 1, &rc, DT_CALCRECT | DT_RIGHT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
						break;
					case ES_LEFT:
						targetDC.DrawText(&pBuffer[copied - 1], 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
						break;
				}
				textRectangle.right += rc.Width();

				targetDC.SelectFont(hPreviousFont);
				ReleaseDC(targetDC);
			}

			*pValue = BOOL2VARIANTBOOL((textRectangle.left < clientRectangle.left) || (textRectangle.right > clientRectangle.right));
		}
		delete[] pBuffer;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, LONG draggedTextFirstChar/* = -1*/, LONG draggedTextLastChar/* = -1*/, LONG itemCountToDisplay/* = 0*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSERT(draggedTextFirstChar >= 0);
		ATLASSERT(draggedTextLastChar >= 0);
		if(draggedTextFirstChar < 0 || draggedTextLastChar < 0) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = m_hWnd;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	dragDropStatus.draggedTextFirstChar = draggedTextFirstChar;
	dragDropStatus.draggedTextLastChar = draggedTextLastChar;
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is TextBox[U|A], so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	if(pOLEDataObject) {
		Raise_OLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	KillTimer(timers.ID_DRAGSCROLL);
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_OLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP TextBox::PositionToCharIndex(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG* pCharacterIndex/* = NULL*/, LONG* pLineIndex/* = NULL*/)
{
	if(IsWindow()) {
		if(pCharacterIndex) {
			*pCharacterIndex = -1;
		}
		if(pLineIndex) {
			*pLineIndex = -1;
		}
		LRESULT lr = SendMessage(EM_CHARFROMPOS, 0, MAKELPARAM(x, y));
		if((LOWORD(lr) != 65535) || (HIWORD(lr) != 65535)) {
			if(pCharacterIndex) {
				*pCharacterIndex = LOWORD(lr);
			}
			if(pLineIndex) {
				*pLineIndex = HIWORD(lr);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::Refresh(void)
{
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		Invalidate();
		UpdateWindow();
		dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}

STDMETHODIMP TextBox::ReplaceSelectedText(BSTR replacementText, VARIANT_BOOL undoable/* = VARIANT_FALSE*/)
{
	if(IsWindow()) {
		SendMessage(EM_REPLACESEL, VARIANTBOOL2BOOL(undoable), reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(COLE2CT(replacementText))));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP TextBox::Scroll(ScrollAxisConstants axis, ScrollDirectionConstants directionAndIntensity, LONG linesToScrollVertically/* = 0*/, LONG charactersToScrollHorizontally/* = 0*/)
{
	if(IsWindow()) {
		if(directionAndIntensity == sdCustom) {
			SendMessage(EM_LINESCROLL, ((axis & saHorizontal) == saHorizontal) ? charactersToScrollHorizontally : 0, ((axis & saVertical) == saVertical) ? linesToScrollVertically : 0);
		} else {
			if(axis & saVertical) {
				SendMessage(EM_SCROLL, directionAndIntensity, 0);
			}
			if(axis & saHorizontal) {
				SendMessage(WM_HSCROLL, directionAndIntensity, 0);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::ScrollCaretIntoView(void)
{
	if(IsWindow()) {
		SendMessage(EM_SCROLLCARET, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, LONG characterIndex)
{
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		switch(relativePosition) {
			case impNowhere:
				if(SendMessage(EM_SETINSERTMARK, FALSE, -1)) {
					hr = S_OK;
				}
				break;
			case impBefore:
				if(SendMessage(EM_SETINSERTMARK, FALSE, characterIndex)) {
					hr = S_OK;
				}
				break;
			case impAfter:
				if(SendMessage(EM_SETINSERTMARK, TRUE, characterIndex)) {
					hr = S_OK;
				}
				break;
			case impDontChange:
				hr = S_OK;
				break;
		}
		dragDropStatus.ShowDragImage(TRUE);
	}

	return hr;
}

STDMETHODIMP TextBox::SetSelection(LONG selectionStart, LONG selectionEnd)
{
	if(IsWindow()) {
		SendMessage(EM_SETSEL, selectionStart, selectionEnd);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::ShowBalloonTip(BSTR title, BSTR text, BalloonTipIconConstants icon/* = btiNone*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		EDITBALLOONTIP balloonTipDetails = {0};
		balloonTipDetails.cbStruct = sizeof(EDITBALLOONTIP);
		balloonTipDetails.pszText = OLE2W(text);
		balloonTipDetails.pszTitle = OLE2W(title);
		balloonTipDetails.ttiIcon = icon;
		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(EM_SHOWBALLOONTIP, 0, reinterpret_cast<LPARAM>(&balloonTipDetails)));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextBox::Undo(VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(EM_UNDO, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}


LRESULT TextBox::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT TextBox::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y, &showDefaultMenu);
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT TextBox::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT TextBox::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnEraseBkGnd(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(properties.useCustomFormattingRectangle) {
		SendMessage(EM_SETRECTNP, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	return lr;
}

LRESULT TextBox::OnGetDlgCode(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(properties.acceptTabKey) {
		lr |= DLGC_WANTTAB;
	} else {
		lr &= ~DLGC_WANTTAB;
	}
	return lr;
}

LRESULT TextBox::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if((ime != imeNoControl) && (GetFocus() == *this)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(*this, ime);
	}
	return lr;
}

LRESULT TextBox::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TextBox::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT TextBox::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<TextBox>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT TextBox::OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TextBox::OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = 0;
	if(properties.allowDragDrop) {
		int selectionStart = 0;
		int selectionEnd = 0;
		SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
		if(abs(selectionEnd - selectionStart) > 0) {
			// we've a selection - check whether the mouse cursor is over selected text
			if(IsOverSelectedText(lParam)) {
				dragDropStatus.candidate.textFirstChar = selectionStart;
				dragDropStatus.candidate.textLastChar = selectionEnd - 1;
				dragDropStatus.candidate.button = MK_LBUTTON;
				dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
				dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);

				dragDropStatus.candidate.bufferedLButtonDown.wParam = wParam;
				dragDropStatus.candidate.bufferedLButtonDown.lParam = lParam;

				SetCapture();
			} else {
				lr = DefWindowProc(message, wParam, lParam);
			}
		} else {
			// there's no selected text
			lr = DefWindowProc(message, wParam, lParam);
		}
	} else {
		lr = DefWindowProc(message, wParam, lParam);
	}

	return lr;
}

LRESULT TextBox::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.allowDragDrop) {
		if(dragDropStatus.candidate.button == MK_LBUTTON) {
			if((dragDropStatus.candidate.bufferedLButtonDown.wParam != 0) && (dragDropStatus.candidate.bufferedLButtonDown.lParam != 0)) {
				DefWindowProc(WM_LBUTTONDOWN, dragDropStatus.candidate.bufferedLButtonDown.wParam, dragDropStatus.candidate.bufferedLButtonDown.lParam);
			}
			dragDropStatus.candidate.Reset();
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TextBox::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TextBox::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TextBox::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TextBox::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT TextBox::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TextBox::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT TextBox::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	if(properties.allowDragDrop) {
		if((dragDropStatus.candidate.textFirstChar != -1) && (dragDropStatus.candidate.textLastChar != -1)) {
			// calculate the rectangle, that the mouse cursor must leave to start dragging
			int clickRectWidth = GetSystemMetrics(SM_CXDRAG);
			if(clickRectWidth < 4) {
				clickRectWidth = 4;
			}
			int clickRectHeight = GetSystemMetrics(SM_CYDRAG);
			if(clickRectHeight < 4) {
				clickRectHeight = 4;
			}
			CRect rc(dragDropStatus.candidate.position.x - clickRectWidth, dragDropStatus.candidate.position.y - clickRectHeight, dragDropStatus.candidate.position.x + clickRectWidth, dragDropStatus.candidate.position.y + clickRectHeight);

			if(!rc.PtInRect(CPoint(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)))) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				switch(dragDropStatus.candidate.button) {
					case MK_LBUTTON:
						Raise_BeginDrag(dragDropStatus.candidate.textFirstChar, dragDropStatus.candidate.textLastChar, button, shift, dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y);
						break;
					case MK_RBUTTON:
						Raise_BeginRDrag(dragDropStatus.candidate.textFirstChar, dragDropStatus.candidate.textLastChar, button, shift, dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y);
						break;
				}
				dragDropStatus.candidate.textFirstChar = -1;
				dragDropStatus.candidate.textLastChar = -1;
			}
		}
	}

	if(dragDropStatus.IsDragging()) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_DragMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TextBox::OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT TextBox::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(message == WM_PAINT) {
		HDC hDC = GetDC();
		DrawInsertionMark(hDC);
		ReleaseDC(hDC);
	} else {
		DrawInsertionMark(reinterpret_cast<HDC>(wParam));
	}
	return lr;
}

LRESULT TextBox::OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TextBox::OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = 0;
	if(properties.allowDragDrop) {
		int selectionStart = 0;
		int selectionEnd = 0;
		SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
		if(abs(selectionEnd - selectionStart) > 0) {
			// we've a selection - check whether the mouse cursor is over selected text
			if(IsOverSelectedText(lParam)) {
				dragDropStatus.candidate.textFirstChar = selectionStart;
				dragDropStatus.candidate.textLastChar = selectionEnd - 1;
				dragDropStatus.candidate.button = MK_RBUTTON;
				dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
				dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);

				dragDropStatus.candidate.bufferedRButtonDown.wParam = wParam;
				dragDropStatus.candidate.bufferedRButtonDown.lParam = lParam;

				SetCapture();
			} else {
				lr = DefWindowProc(message, wParam, lParam);
			}
		} else {
			// there's no selected text
			lr = DefWindowProc(message, wParam, lParam);
		}
	} else {
		lr = DefWindowProc(message, wParam, lParam);
	}

	return lr;
}

LRESULT TextBox::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.allowDragDrop) {
		if(dragDropStatus.candidate.button == MK_RBUTTON) {
			if((dragDropStatus.candidate.bufferedRButtonDown.wParam != 0) && (dragDropStatus.candidate.bufferedRButtonDown.lParam != 0)) {
				DefWindowProc(WM_RBUTTONDOWN, dragDropStatus.candidate.bufferedRButtonDown.wParam, dragDropStatus.candidate.bufferedRButtonDown.lParam);
			}
			dragDropStatus.candidate.Reset();
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TextBox::OnScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(insertMark.characterIndex != -1) {
		// remove the insertion mark - we would get drawing glitches otherwise
		CRect oldInsertMarkRect;
		// calculate the current insertion mark's rectangle
		LRESULT lr = SendMessage(EM_POSFROMCHAR, insertMark.characterIndex, 0);
		if(lr == -1) {
			if(insertMark.characterIndex == 0 && GetWindowTextLength() == 0) {
				// HACK: If the control is empty, retrieving the first character's position (of course) fails.
				lr = MAKELRESULT(1, 1);
			}
		}
		if(lr != -1) {
			CDCHandle targetDC = GetDC();
			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			HFONT hPreviousFont = targetDC.SelectFont(hFont);

			if(insertMark.afterChar) {
				int bufferSize = GetWindowTextLength() + 1;
				LPTSTR pBuffer = new TCHAR[bufferSize];
				GetWindowText(pBuffer, bufferSize);

				// calculate character width
				CRect rc;
				targetDC.DrawText(&pBuffer[insertMark.characterIndex], 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
				oldInsertMarkRect.left = LOWORD(lr) + rc.Width();

				delete[] pBuffer;
			} else {
				oldInsertMarkRect.left = LOWORD(lr) - 1;
			}
			oldInsertMarkRect.right = oldInsertMarkRect.left + 2;
			oldInsertMarkRect.top = HIWORD(lr);

			LONG lineHeight = 0;
			get_LineHeight(&lineHeight);
			oldInsertMarkRect.bottom = oldInsertMarkRect.top + lineHeight;

			targetDC.SelectFont(hPreviousFont);
			ReleaseDC(targetDC);

			oldInsertMarkRect.InflateRect(1, 1);
		}

		// redraw
		if(oldInsertMarkRect.Width() > 0) {
			InvalidateRect(&oldInsertMarkRect);
		}

		insertMark.hidden = TRUE;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(insertMark.characterIndex != -1) {
		// now draw the insertion mark again
		insertMark.hidden = FALSE;
		HDC hDC = GetDC();
		DrawInsertionMark(hDC);
		ReleaseDC(hDC);
	}

	if(!(properties.disabledEvents & deScrolling)) {
		if((LOWORD(wParam) == SB_THUMBPOSITION) || (LOWORD(wParam) == SB_THUMBTRACK)) {
			switch(message) {
				case WM_HSCROLL:
					Raise_Scrolling(saHorizontal);
					break;
				case WM_VSCROLL:
					Raise_Scrolling(saVertical);
					break;
			}
		}
	}

	return lr;
}

LRESULT TextBox::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		ScreenToClient(&mousePosition);
		if(IsOverSelectedText(MAKELPARAM(mousePosition.x, mousePosition.y))) {
			if(properties.selectedTextMousePointer == mpCustom) {
				if(properties.selectedTextMouseIcon.pPictureDisp) {
					CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.selectedTextMouseIcon.pPictureDisp);
					if(pPicture) {
						OLE_HANDLE h = NULL;
						pPicture->get_Handle(&h);
						hCursor = static_cast<HCURSOR>(LongToHandle(h));
					}
				}
			} else {
				hCursor = MousePointerConst2hCursor(properties.selectedTextMousePointer);
			}
		} else {
			if(properties.mousePointer == mpCustom) {
				if(properties.mouseIcon.pPictureDisp) {
					CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
					if(pPicture) {
						OLE_HANDLE h = NULL;
						pPicture->get_Handle(&h);
						hCursor = static_cast<HCURSOR>(LongToHandle(h));
					}
				}
			} else {
				hCursor = MousePointerConst2hCursor(properties.mousePointer);
			}
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT TextBox::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<TextBox>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<TextBox>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT TextBox::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TXTBOX_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_TXTBOX_FONT);
	}

	return lr;
}

LRESULT TextBox::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT TextBox::OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TXTBOX_TEXT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		if(!(properties.disabledEvents & deTextChangedEvents)) {
			if(GetStyle() & ES_MULTILINE) {
				Raise_TextChanged();
			}
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_TEXT);
	SendOnDataChange();
	return lr;
}

LRESULT TextBox::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_DRAGSCROLL:
			AutoScroll();
			break;

		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT TextBox::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		Invalidate();
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT TextBox::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TextBox::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT TextBox::OnReflectedCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
	SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.foreColor));
	if(!(properties.backColor & 0x80000000)) {
		SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
		return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
	} else {
		return reinterpret_cast<LRESULT>(GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
	}
}

LRESULT TextBox::OnReflectedCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(properties.disabledForeColor != static_cast<OLE_COLOR>(-1)) {
		SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.disabledForeColor));
	}

	if(RunTimeHelper::IsCommCtrl6() && properties.disabledBackColor == static_cast<OLE_COLOR>(-1)) {
		if(flags.usingThemes) {
			RECT clientRectangle;
			::GetClientRect(reinterpret_cast<HWND>(lParam), &clientRectangle);
			int state = SaveDC(reinterpret_cast<HDC>(wParam));				// if we don't do this, we get a large font if the control is read-only and sits inside a VB6 frame
			if(RunTimeHelper::IsVista() && !(GetStyle() & ES_MULTILINE)) {
				/* NOTE: This solution produces more flickering then using a pattern brush, but there have been reports that
				 *       users get a black background if the control is very large and we use a pattern brush.
				 *       Also don't pass DTPB_USECTLCOLORSTATIC here - it will make the text vanish in the Events sample (event log).
				 */
				flags.useTransparentTextBackground = (DrawThemeParentBackgroundEx(reinterpret_cast<HWND>(lParam), reinterpret_cast<HDC>(wParam), 0, &clientRectangle) == S_OK);
			} else {
				// Problem: If we return a NULL_BRUSH, no text will be drawn, if the control is read-only.
				#if TRUE
					CDC memoryDC;
					memoryDC.CreateCompatibleDC(reinterpret_cast<HDC>(wParam));
					CBitmap memoryBitmap;
					memoryBitmap.CreateCompatibleBitmap(reinterpret_cast<HDC>(wParam), clientRectangle.right - clientRectangle.left, clientRectangle.bottom - clientRectangle.top);
					HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(memoryBitmap);
					memoryDC.SetViewportOrg(-clientRectangle.left, -clientRectangle.top);

					memoryDC.FillRect(&clientRectangle, GetSysColorBrush(COLOR_BTNFACE));					// if we don't do this, we get a black background if the control is read-only and inside a VB6 frame
					flags.useTransparentTextBackground = (DrawThemeParentBackground(reinterpret_cast<HWND>(lParam), memoryDC, &clientRectangle) == S_OK);

					memoryDC.SelectBitmap(hPreviousBitmap);
					if(!themedBackBrush.IsNull()) {
						themedBackBrush.DeleteObject();
					}
					themedBackBrush.CreatePatternBrush(memoryBitmap);
				#else
					FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, GetSysColorBrush(COLOR_BTNFACE));					// if we don't do this, we get a black background if the control is read-only and inside a VB6 frame
					flags.useTransparentTextBackground = (DrawThemeParentBackground(reinterpret_cast<HWND>(lParam), reinterpret_cast<HDC>(wParam), &clientRectangle) == S_OK);
				#endif
			}
			RestoreDC(reinterpret_cast<HDC>(wParam), state);
		} else {
			flags.useTransparentTextBackground = FALSE;
		}
		if(flags.useTransparentTextBackground) {
			SetBkMode(reinterpret_cast<HDC>(wParam), TRANSPARENT);
			if(themedBackBrush.IsNull()) {
				return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(NULL_BRUSH)));
			} else {
				return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(themedBackBrush));
			}
		}
	}

	if(properties.disabledBackColor == static_cast<OLE_COLOR>(-1)) {
		SetBkColor(reinterpret_cast<HDC>(wParam), GetSysColor(COLOR_3DFACE));
		return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_3DFACE));
	} else if(!(properties.disabledBackColor & 0x80000000)) {
		SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.disabledBackColor));
		SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.disabledBackColor));
		return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
	} else {
		SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.disabledBackColor));
		return reinterpret_cast<LRESULT>(GetSysColorBrush(properties.disabledBackColor & 0x0FFFFFFF));
	}
}

LRESULT TextBox::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.draggedTextFirstChar != -1 && dragDropStatus.draggedTextLastChar != -1) {
		if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = TRUE; //CreateVistaOLEDragImage(dragDropStatus.draggedTextFirstChar, dragDropStatus.draggedTextLastChar, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEDragImage(dragDropStatus.draggedTextFirstChar, dragDropStatus.draggedTextLastChar, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT TextBox::OnFmtLines(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	properties.insertSoftLineBreaks = (lr != 0);
	return lr;
}

LRESULT TextBox::OnGetInsertMark(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	ATLASSERT(lParam);
	if(!lParam) {
		return FALSE;
	}

	LPECINSERTMARK pDetails = reinterpret_cast<LPECINSERTMARK>(lParam);
	if(pDetails->size == sizeof(ECINSERTMARK)) {
		pDetails->characterIndex = insertMark.characterIndex;
		pDetails->afterChar = insertMark.afterChar;
		return TRUE;
	}
	return FALSE;
}

LRESULT TextBox::OnGetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	return insertMark.color;
}

LRESULT TextBox::OnSetCueBanner(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TXTBOX_CUEBANNER) == S_FALSE) {
		return FALSE;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.displayCueBannerOnFocus = static_cast<BOOL>(wParam);
		properties.cueBanner = reinterpret_cast<LPWSTR>(lParam);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_CUEBANNER);
	SendOnDataChange();
	return lr;
}

LRESULT TextBox::OnSetInsertMark(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	CRect oldInsertMarkRect;
	CRect newInsertMarkRect;

	if(insertMark.characterIndex != -1) {
		// calculate the current insertion mark's rectangle
		LRESULT lr = SendMessage(EM_POSFROMCHAR, insertMark.characterIndex, 0);
		if(lr == -1) {
			if(insertMark.characterIndex == 0 && GetWindowTextLength() == 0) {
				// HACK: If the control is empty, retrieving the first character's position (of course) fails.
				lr = MAKELRESULT(1, 1);
			}
		}
		if(lr != -1) {
			CDCHandle targetDC = GetDC();
			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			HFONT hPreviousFont = targetDC.SelectFont(hFont);

			if(insertMark.afterChar) {
				int bufferSize = GetWindowTextLength() + 1;
				LPTSTR pBuffer = new TCHAR[bufferSize];
				GetWindowText(pBuffer, bufferSize);

				// calculate character width
				CRect rc;
				targetDC.DrawText(&pBuffer[insertMark.characterIndex], 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
				oldInsertMarkRect.left = LOWORD(lr) + rc.Width();

				delete[] pBuffer;
			} else {
				oldInsertMarkRect.left = LOWORD(lr) - 1;
			}
			oldInsertMarkRect.right = oldInsertMarkRect.left + 2;
			oldInsertMarkRect.top = HIWORD(lr);

			LONG lineHeight = 0;
			get_LineHeight(&lineHeight);
			oldInsertMarkRect.bottom = oldInsertMarkRect.top + lineHeight;

			targetDC.SelectFont(hPreviousFont);
			ReleaseDC(targetDC);

			oldInsertMarkRect.InflateRect(1, 1);
		}
	}

	insertMark.afterChar = static_cast<BOOL>(wParam);
	insertMark.characterIndex = static_cast<int>(lParam);

	if(insertMark.characterIndex != -1) {
		// calculate the current insertion mark's rectangle
		LRESULT lr = SendMessage(EM_POSFROMCHAR, insertMark.characterIndex, 0);
		if(lr == -1) {
			if(insertMark.characterIndex == 0 && GetWindowTextLength() == 0) {
				// HACK: If the control is empty, retrieving the first character's position (of course) fails.
				lr = MAKELRESULT(1, 1);
			}
		}
		if(lr != -1) {
			CDCHandle targetDC = GetDC();
			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			HFONT hPreviousFont = targetDC.SelectFont(hFont);

			if(insertMark.afterChar) {
				int bufferSize = GetWindowTextLength() + 1;
				LPTSTR pBuffer = new TCHAR[bufferSize];
				GetWindowText(pBuffer, bufferSize);

				// calculate character width
				CRect rc;
				targetDC.DrawText(&pBuffer[insertMark.characterIndex], 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
				newInsertMarkRect.left = LOWORD(lr) + rc.Width();

				delete[] pBuffer;
			} else {
				newInsertMarkRect.left = LOWORD(lr) - 1;
			}
			newInsertMarkRect.right = newInsertMarkRect.left + 2;
			newInsertMarkRect.top = HIWORD(lr);

			LONG lineHeight = 0;
			get_LineHeight(&lineHeight);
			newInsertMarkRect.bottom = newInsertMarkRect.top + lineHeight;

			targetDC.SelectFont(hPreviousFont);
			ReleaseDC(targetDC);

			newInsertMarkRect.InflateRect(1, 1);
		}
	}

	if(newInsertMarkRect != oldInsertMarkRect) {
		// redraw
		if(oldInsertMarkRect.Width() > 0) {
			InvalidateRect(&oldInsertMarkRect);
		}
		if(newInsertMarkRect.Width() > 0) {
			InvalidateRect(&newInsertMarkRect);
		}
	}

	// always succeed
	return TRUE;
}

LRESULT TextBox::OnSetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	COLORREF previousColor = insertMark.color;
	insertMark.color = static_cast<COLORREF>(lParam);

	// calculate insertion mark rectangle
	CRect insertMarkRect;
	if(insertMark.characterIndex != -1) {
		// calculate the current insertion mark's rectangle
		LRESULT lr = SendMessage(EM_POSFROMCHAR, insertMark.characterIndex, 0);
		if(lr == -1) {
			if(insertMark.characterIndex == 0 && GetWindowTextLength() == 0) {
				// HACK: If the control is empty, retrieving the first character's position (of course) fails.
				lr = MAKELRESULT(1, 1);
			}
		}
		if(lr != -1) {
			CDCHandle targetDC = GetDC();
			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			HFONT hPreviousFont = targetDC.SelectFont(hFont);

			if(insertMark.afterChar) {
				int bufferSize = GetWindowTextLength() + 1;
				LPTSTR pBuffer = new TCHAR[bufferSize];
				GetWindowText(pBuffer, bufferSize);

				// calculate character width
				CRect rc;
				targetDC.DrawText(&pBuffer[insertMark.characterIndex], 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
				insertMarkRect.left = LOWORD(lr) + rc.Width();

				delete[] pBuffer;
			} else {
				insertMarkRect.left = LOWORD(lr) - 1;
			}
			insertMarkRect.right = insertMarkRect.left + 2;
			insertMarkRect.top = HIWORD(lr);

			LONG lineHeight = 0;
			get_LineHeight(&lineHeight);
			insertMarkRect.bottom = insertMarkRect.top + lineHeight;

			targetDC.SelectFont(hPreviousFont);
			ReleaseDC(targetDC);

			insertMarkRect.InflateRect(1, 1);
		}

		// redraw this rectangle
		InvalidateRect(&insertMarkRect);
	}

	return previousColor;
}

LRESULT TextBox::OnSetTabStops(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr && !IsInDesignMode()) {
		#ifdef USE_STL
			properties.tabStops.clear();
		#else
			properties.tabStops.RemoveAll();
		#endif
		if(wParam == 0) {
			properties.tabWidthInDTUs = 32;

			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			CDC memoryDC;
			memoryDC.CreateCompatibleDC();
			HFONT hPreviousFont = memoryDC.SelectFont(hFont);
			SIZE textExtent = {0};
			memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
			int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
			memoryDC.SelectFont(hPreviousFont);
			int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

			// convert to pixels
			properties.tabWidthInPixels = 2 * properties.tabWidthInDTUs * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;
		} else if(wParam == 1) {
			properties.tabWidthInDTUs = *reinterpret_cast<PUINT>(lParam);
			#ifdef USE_STL
				properties.tabStops.push_back(properties.tabWidthInDTUs);
			#else
				properties.tabStops.Add(properties.tabWidthInDTUs);
			#endif

			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			CDC memoryDC;
			memoryDC.CreateCompatibleDC();
			HFONT hPreviousFont = memoryDC.SelectFont(hFont);
			SIZE textExtent = {0};
			memoryDC.GetTextExtent(TEXT("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &textExtent);
			int averageCharWidthOfUsedFont = (textExtent.cx / 26 + 1) / 2;
			memoryDC.SelectFont(hPreviousFont);
			int averageCharWidthOfSystemFont = LOWORD(GetDialogBaseUnits());

			// convert to pixels
			properties.tabWidthInPixels = 2 * properties.tabWidthInDTUs * averageCharWidthOfUsedFont / averageCharWidthOfSystemFont;
		} else {
			for(UINT i = 0; i < wParam; ++i) {
				#ifdef USE_STL
					properties.tabStops.push_back(reinterpret_cast<PUINT>(lParam)[i]);
				#else
					properties.tabStops.Add(reinterpret_cast<PUINT>(lParam)[i]);
				#endif
			}
		}
	}
	return lr;
}


LRESULT TextBox::OnReflectedAlign(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	switch(notifyCode) {
		case EN_ALIGN_LTR_EC:
			Raise_WritingDirectionChanged(wdLeftToRight);
			break;
		case EN_ALIGN_RTL_EC:
			Raise_WritingDirectionChanged(wdRightToLeft);
			break;
	}
	return 0;
}

LRESULT TextBox::OnReflectedChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(!(properties.disabledEvents & deTextChangedEvents)) {
		Raise_TextChanged();
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_TXTBOX_TEXT);
	SendOnDataChange();

	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnReflectedErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_OutOfMemory();
	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnReflectedMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_TruncatedText();
	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnReflectedScroll(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deScrolling)) {
		switch(notifyCode) {
			case EN_HSCROLL:
				Raise_Scrolling(saHorizontal);
				break;
			case EN_VSCROLL:
				Raise_Scrolling(saVertical);
				break;
		}
	}

	return 0;
}

LRESULT TextBox::OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(*this);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(*this, ime);
			}
		}
	}

	if(SendMessage(WM_GETDLGCODE, 0, 0) & DLGC_HASSETSEL) {
		if(!properties.acceptTabKey) {
			if(GetAsyncKeyState(VK_TAB) & 0x8000) {
				SendMessage(EM_SETSEL, 0, -1);
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT TextBox::OnReflectedUpdate(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deBeforeDrawText)) {
		Raise_BeforeDrawText();
	}

	return 0;
}


void TextBox::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT TextBox::Raise_AbortedDrag(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AbortedDrag();
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_BeforeDrawText(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeforeDrawText();
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_BeginDrag(LONG firstChar, LONG lastChar, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginDrag(firstChar, lastChar, button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_BeginRDrag(LONG firstChar, LONG lastChar, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginRDrag(firstChar, lastChar, button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// propose the middle of the control's client rectangle as the menu's position
				CRect clientRectangle;
				GetClientRect(&clientRectangle);
				CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
			} else {
				return S_OK;
			}
		}

		return Fire_ContextMenu(button, shift, x, y, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(dragDropStatus.hDragImageList) {
		DWORD position = GetMessagePos();
		ImageList_DragMove(GET_X_LPARAM(position), GET_Y_LPARAM(position));
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePosition(x, y);
		CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePosition);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePosition);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePosition.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePosition.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePosition.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePosition.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_DragMouseMove(button, shift, x, y, &autoHScrollVelocity, &autoVScrollVelocity);
	//}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT TextBox::Raise_Drop(void)
{
	//if(m_nFreezeEvents == 0) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		return Fire_Drop(button, shift, mousePosition.x, mousePosition.y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();

		return Fire_MouseDown(button, shift, x, y);
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();
		return S_OK;
	}
}

inline HRESULT TextBox::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT TextBox::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.LeaveControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT TextBox::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y);
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(GetCapture() == *this) {
			ReleaseCapture();
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT TextBox::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(button, shift, x, y, scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
				always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT TextBox::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT TextBox::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_OLEDragLeave(void)
{
	KillTimer(timers.ID_DRAGSCROLL);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT TextBox::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		CPoint mousePos(mousePosition.x, mousePosition.y);
		CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT TextBox::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_OLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT TextBox::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT TextBox::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_OutOfMemory(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OutOfMemory();
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_Scrolling(ScrollAxisConstants axis)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Scrolling(axis);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_TextChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TextChanged();
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_TruncatedText(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TruncatedText();
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_WritingDirectionChanged(WritingDirectionConstants newWritingDirection)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_WritingDirectionChanged(newWritingDirection);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT TextBox::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y);
	//}
	//return S_OK;
}


void TextBox::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		if(IsWindow()) {
			GetWindowText(&properties.text);
		}

		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

IMEModeConstants TextBox::GetCurrentIMEContextMode(HWND hWnd)
{
	IMEModeConstants imeContextMode = imeNoControl;

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		imeContextMode = imeInherit;
	} else {
		HIMC hIMC = ImmGetContext(hWnd);
		if(hIMC) {
			if(ImmGetOpenStatus(hIMC)) {
				DWORD conversionMode = 0;
				DWORD sentenceMode = 0;
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				if(conversionMode & IME_CMODE_NATIVE) {
					if(conversionMode & IME_CMODE_KATAKANA) {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeKatakanaHalf];
						} else {
							imeContextMode = countryTable[imeAlphaFull];
						}
					} else {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeHiragana];
						} else {
							imeContextMode = countryTable[imeKatakana];
						}
					}
				} else {
					if(conversionMode & IME_CMODE_FULLSHAPE) {
						imeContextMode = countryTable[imeAlpha];
					} else {
						imeContextMode = countryTable[imeHangulFull];
					}
				}
			} else {
				imeContextMode = countryTable[imeDisable];
			}
			ImmReleaseContext(hWnd, hIMC);
		} else {
			imeContextMode = countryTable[imeOn];
		}
	}
	return imeContextMode;
}

void TextBox::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
{
	if((IMEMode == imeInherit) || (IMEMode == imeNoControl) || !::IsWindow(hWnd)) {
		return;
	}

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		return;
	}

	// update IME mode
	HIMC hIMC = ImmGetContext(hWnd);
	if(IMEMode == imeDisable) {
		// disable IME
		if(hIMC) {
			// close the IME
			if(ImmGetOpenStatus(hIMC)) {
				ImmSetOpenStatus(hIMC, FALSE);
			}
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
		// remove the control's association to the IME context
		HIMC h = ImmAssociateContext(hWnd, NULL);
		if(h) {
			IMEFlags.hDefaultIMC = h;
		}
		return;
	} else {
		// enable IME
		if(!hIMC) {
			if(!IMEFlags.hDefaultIMC) {
				// create an IME context
				hIMC = ImmCreateContext();
				if(hIMC) {
					// associate the control with the IME context
					ImmAssociateContext(hWnd, hIMC);
				}
			} else {
				// associate the control with the default IME context
				ImmAssociateContext(hWnd, IMEFlags.hDefaultIMC);
			}
		} else {
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
	}

	hIMC = ImmGetContext(hWnd);
	if(hIMC) {
		DWORD conversionMode = 0;
		DWORD sentenceMode = 0;
		switch(IMEMode) {
			case imeOn:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				break;
			case imeOff:
				// close IME
				ImmSetOpenStatus(hIMC, FALSE);
				break;
			default:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				// switch conversion
				switch(IMEMode) {
					case imeHiragana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_KATAKANA;
						break;
					case imeKatakana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeKatakanaHalf:
						conversionMode |= (IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
					case imeAlphaFull:
						conversionMode |= IME_CMODE_FULLSHAPE;
						conversionMode &= ~(IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeAlpha:
						conversionMode |= IME_CMODE_ALPHANUMERIC;
						conversionMode &= ~(IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeHangulFull:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeHangul:
						conversionMode |= IME_CMODE_NATIVE;
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
				}
				ImmSetConversionStatus(hIMC, conversionMode, sentenceMode);
				break;
		}
		// each ImmGetContext() needs a ImmReleaseContext()
		ImmReleaseContext(hWnd, hIMC);
		hIMC = NULL;
	}
}

IMEModeConstants TextBox::GetEffectiveIMEMode(void)
{
	IMEModeConstants IMEMode = properties.IMEMode;
	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

DWORD TextBox::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD TextBox::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | /*WS_CLIPCHILDREN | */WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}
	if(properties.scrollBars & sbVertical) {
		style |= WS_VSCROLL;
	}
	if(properties.scrollBars & sbHorizontal) {
		style |= WS_HSCROLL;
	}

	if(properties.acceptNumbersOnly) {
		style |= ES_NUMBER;
	}
	if(properties.alwaysShowSelection) {
		style |= ES_NOHIDESEL;
	}
	if(properties.autoScrolling & asVertical) {
		style |= ES_AUTOVSCROLL;
	}
	if(properties.autoScrolling & asHorizontal) {
		style |= ES_AUTOHSCROLL;
	}
	switch(properties.characterConversion) {
		case ccLowerCase:
			style |= ES_LOWERCASE;
			break;
		case ccUpperCase:
			style |= ES_UPPERCASE;
			break;
	}
	if(properties.doOEMConversion) {
		style |= ES_OEMCONVERT;
	}
	switch(properties.hAlignment) {
		case halLeft:
			style |= ES_LEFT;
			break;
		case halCenter:
			style |= ES_CENTER;
			break;
		case halRight:
			style |= ES_RIGHT;
			break;
	}
	if(properties.multiLine) {
		style |= ES_MULTILINE;
	}
	return style;
}

void TextBox::SendConfigurationMessages(void)
{
	DWORD flags = static_cast<DWORD>(SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0));
	if(properties.cancelIMECompositionOnSetFocus) {
		flags |= EIMES_CANCELCOMPSTRINFOCUS;
	} else {
		flags &= ~EIMES_CANCELCOMPSTRINFOCUS;
	}
	if(properties.completeIMECompositionOnKillFocus) {
		flags |= EIMES_COMPLETECOMPSTRKILLFOCUS;
	} else {
		flags &= ~EIMES_COMPLETECOMPSTRKILLFOCUS;
	}
	SendMessage(EM_SETIMESTATUS, EMSIS_COMPOSITIONSTRING, flags);
	SendMessage(EM_FMTLINES, properties.insertSoftLineBreaks, 0);
	SendMessage(EM_SETLIMITTEXT, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength), 0);
	SendMessage(EM_SETREADONLY, properties.readOnly, 0);
	if(properties.usePasswordChar) {
		if(properties.passwordChar == 0) {
			#ifdef UNICODE
				if(RunTimeHelper::IsCommCtrl6()) {
					SendMessage(EM_SETPASSWORDCHAR, L'\x25CF', 0);
				} else {
					SendMessage(EM_SETPASSWORDCHAR, L'*', 0);
				}
			#else
				SendMessage(EM_SETPASSWORDCHAR, '*', 0);
			#endif
		} else {
			SendMessage(EM_SETPASSWORDCHAR, properties.passwordChar, 0);
		}
	} else {
		SendMessage(EM_SETPASSWORDCHAR, 0, 0);
	}

	#ifdef USE_STL
		if(properties.tabStops.empty()) {
	#else
		if(properties.tabStops.IsEmpty()) {
	#endif
		if(properties.tabWidthInPixels == -1) {
			SendMessage(EM_SETTABSTOPS, 0, 0);
		} else {
			UINT distance = properties.tabWidthInDTUs;
			SendMessage(EM_SETTABSTOPS, 1, reinterpret_cast<LPARAM>(&distance));
		}
	} else {
		#ifdef USE_STL
			size_t numberOfStops = properties.tabStops.size();
		#else
			size_t numberOfStops = properties.tabStops.GetCount();
		#endif
		PUINT pTabStops = new UINT[numberOfStops];
		for(size_t i = 0; i < numberOfStops; ++i) {
			pTabStops[i] = properties.tabStops[i];
		}
		SendMessage(EM_SETTABSTOPS, numberOfStops, reinterpret_cast<LPARAM>(pTabStops));
		delete[] pTabStops;
	}

	if(properties.useCustomFormattingRectangle) {
		SendMessage(EM_SETRECTNP, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	} else {
		SendMessage(EM_SETRECTNP, 0, NULL);
	}

	SetWindowText(COLE2CT(properties.text));
	if(RunTimeHelper::IsCommCtrl6()) {
		SendMessage(EM_SETCUEBANNER, properties.displayCueBannerOnFocus, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
	}
	SendMessage(EM_SETMODIFY, properties.modified, 0);

	SendMessage(EM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
	ApplyFont();
	SendMessage(EM_SETMARGINS, (EC_LEFTMARGIN | EC_RIGHTMARGIN), MAKELPARAM((properties.leftMargin == -1 ? EC_USEFONTINFO : properties.leftMargin), (properties.rightMargin == -1 ? EC_USEFONTINFO : properties.rightMargin)));
}

HCURSOR TextBox::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

BOOL TextBox::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void TextBox::AutoScroll(void)
{
	LONG realScrollTimeBase = properties.dragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime() / 4;
	}

	if((dragDropStatus.autoScrolling.currentHScrollVelocity != 0) && ((GetStyle() & WS_HSCROLL) == WS_HSCROLL)) {
		if(dragDropStatus.autoScrolling.currentHScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll to the left?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll left
					dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINELEFT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll to the right?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll right
					dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINERIGHT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}

	if((dragDropStatus.autoScrolling.currentVScrollVelocity != 0) && ((GetStyle() & WS_VSCROLL) == WS_VSCROLL)) {
		if(dragDropStatus.autoScrolling.currentVScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll upwardly?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Up) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll up
					dragDropStatus.autoScrolling.lastScroll_Up = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEUP, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll downwards?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Down) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll down
					dragDropStatus.autoScrolling.lastScroll_Down = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEDOWN, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}
}

BOOL TextBox::IsOverSelectedText(LPARAM pt)
{
	BOOL ret = FALSE;

	if(IsWindow()) {
		int selectionStart = 0;
		int selectionEnd = 0;
		SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
		if(abs(selectionEnd - selectionStart) > 0) {
			// we've a selection - check whether the point is over selected text
			LRESULT lr = SendMessage(EM_CHARFROMPOS, 0, pt);
			int charIndexCursor = LOWORD(lr);
			int lineIndexCursor = HIWORD(lr);
			if((charIndexCursor != 65535) || (lineIndexCursor != 65535)) {
				// HACK: EM_LINELENGTH doesn't work as expected, e. g. it may return garbage if ES_CENTER is set
				//WORD cursorLineLength = static_cast<WORD>(SendMessage(EM_LINELENGTH, lineIndexCursor, 0));
				WORD bufferSize = 2048;
				LPTSTR pLineBuffer = new TCHAR[bufferSize + 1];
				*reinterpret_cast<PWORD>(pLineBuffer) = bufferSize;
				WORD cursorLineLength = static_cast<WORD>(SendMessage(EM_GETLINE, lineIndexCursor, reinterpret_cast<LPARAM>(pLineBuffer)));

				if(cursorLineLength > 0) {
					pLineBuffer[cursorLineLength] = TEXT('\0');

					CDCHandle targetDC = GetDC();
					HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
					HFONT hPreviousFont = targetDC.SelectFont(hFont);

					if(lineIndexCursor == static_cast<int>(SendMessage(EM_GETLINECOUNT, 0, 0)) - 1) {
						// check whether we're really in this line using the y coordinate
						CRect rc;
						targetDC.DrawText(TEXT("A"), 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);

						int yNextLine = HIWORD(SendMessage(EM_POSFROMCHAR, charIndexCursor, 0)) + rc.Height();
						if(yNextLine <= GET_Y_LPARAM(pt)) {
							// this will make the next test fail
							++lineIndexCursor;
						}
					}

					int lineOfFirstSelectedChar = static_cast<int>(SendMessage(EM_LINEFROMCHAR, selectionStart, 0));
					int lineOfLastSelectedChar = static_cast<int>(SendMessage(EM_LINEFROMCHAR, selectionEnd - 1, 0));
					if((lineOfFirstSelectedChar <= lineIndexCursor) && (lineIndexCursor <= lineOfLastSelectedChar)) {
						// at least parts of the line the point lies in are selected
						int firstCharInCursorLine = static_cast<int>(SendMessage(EM_LINEINDEX, lineIndexCursor, 0));
						int lastSelectedCharInCursorLine = firstCharInCursorLine + cursorLineLength - 1;
						int firstSelectedCharInCursorLine = max(selectionStart, firstCharInCursorLine);
						lastSelectedCharInCursorLine = min(lastSelectedCharInCursorLine, selectionEnd - 1);

						ATLASSERT(lineIndexCursor == static_cast<int>(SendMessage(EM_LINEFROMCHAR, firstSelectedCharInCursorLine, 0)));
						ATLASSERT(lineIndexCursor == static_cast<int>(SendMessage(EM_LINEFROMCHAR, lastSelectedCharInCursorLine, 0)));
						ATLASSERT(firstCharInCursorLine <= lastSelectedCharInCursorLine);

						// retrieve the left and right border of the selection rectangle in this line
						int leftFirstSelectedCharInLine = static_cast<short>(LOWORD(SendMessage(EM_POSFROMCHAR, firstSelectedCharInCursorLine, 0)));
						int rightLastSelectedCharInLine = static_cast<short>(LOWORD(SendMessage(EM_POSFROMCHAR, lastSelectedCharInCursorLine, 0)));

						// add the last selected character's width
						CRect clientRectangle;
						GetClientRect(&clientRectangle);
						switch(GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
							case ES_CENTER:
								targetDC.DrawText(&pLineBuffer[lastSelectedCharInCursorLine - firstCharInCursorLine], 1, &clientRectangle, DT_CALCRECT | DT_CENTER | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
								break;
							case ES_RIGHT:
								targetDC.DrawText(&pLineBuffer[lastSelectedCharInCursorLine - firstCharInCursorLine], 1, &clientRectangle, DT_CALCRECT | DT_RIGHT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
								break;
							case ES_LEFT:
								targetDC.DrawText(&pLineBuffer[lastSelectedCharInCursorLine - firstCharInCursorLine], 1, &clientRectangle, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
								break;
						}
						rightLastSelectedCharInLine += clientRectangle.Width();

						ret = ((leftFirstSelectedCharInLine <= GET_X_LPARAM(pt)) && (GET_X_LPARAM(pt) <= rightLastSelectedCharInLine));
					}
					targetDC.SelectFont(hPreviousFont);
					ReleaseDC(targetDC);
				}
				delete[] pLineBuffer;
			}
		}
	}
	return ret;
}

void TextBox::DrawInsertionMark(CDCHandle targetDC)
{
	if((insertMark.characterIndex != -1) && !insertMark.hidden) {
		CPen pen;
		pen.CreatePen(PS_SOLID, 2, insertMark.color);
		HPEN hPreviousPen = targetDC.SelectPen(pen);

		CRect insertMarkRect;
		LRESULT lr = SendMessage(EM_POSFROMCHAR, insertMark.characterIndex, 0);
		if(lr == -1) {
			if(insertMark.characterIndex == 0 && GetWindowTextLength() == 0) {
				// HACK: If the control is empty, retrieving the first character's position (of course) fails.
				lr = MAKELRESULT(1, 1);
			}
		}
		if(lr != -1) {
			HFONT hFont = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
			HFONT hPreviousFont = targetDC.SelectFont(hFont);

			if(insertMark.afterChar) {
				int bufferSize = GetWindowTextLength() + 1;
				LPTSTR pBuffer = new TCHAR[bufferSize];
				GetWindowText(pBuffer, bufferSize);

				// calculate character width
				CRect rc;
				targetDC.DrawText(&pBuffer[insertMark.characterIndex], 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
				insertMarkRect.left = LOWORD(lr) + rc.Width();

				delete[] pBuffer;
			} else {
				insertMarkRect.left = LOWORD(lr) - 1;
			}
			insertMarkRect.right = insertMarkRect.left + 2;
			insertMarkRect.top = HIWORD(lr);

			LONG lineHeight = 0;
			get_LineHeight(&lineHeight);
			insertMarkRect.bottom = insertMarkRect.top + lineHeight - 1;

			targetDC.SelectFont(hPreviousFont);
		}

		// draw the line
		targetDC.MoveTo(insertMarkRect.TopLeft());
		targetDC.LineTo(insertMarkRect.left, insertMarkRect.bottom);

		targetDC.SelectPen(hPreviousPen);
	}
}

BOOL TextBox::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL TextBox::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}


HRESULT TextBox::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
{
	if(!hIcon || !pBits || !pWICImagingFactory) {
		return E_FAIL;
	}

	ICONINFO iconInfo;
	GetIconInfo(hIcon, &iconInfo);
	ATLASSERT(iconInfo.hbmColor);
	BITMAP bitmapInfo = {0};
	if(iconInfo.hbmColor) {
		GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bitmapInfo);
	} else if(iconInfo.hbmMask) {
		GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bitmapInfo);
	}
	bitmapInfo.bmHeight = abs(bitmapInfo.bmHeight);
	BOOL needsFrame = ((bitmapInfo.bmWidth < size.cx) || (bitmapInfo.bmHeight < size.cy));
	if(iconInfo.hbmColor) {
		DeleteObject(iconInfo.hbmColor);
	}
	if(iconInfo.hbmMask) {
		DeleteObject(iconInfo.hbmMask);
	}

	HRESULT hr = E_FAIL;

	CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
	if(!needsFrame) {
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapScaler);
	}
	if(needsFrame || SUCCEEDED(hr)) {
		CComPtr<IWICBitmap> pWICBitmapSource = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmapSource);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapSource);
		if(SUCCEEDED(hr)) {
			if(!needsFrame) {
				hr = pWICBitmapScaler->Initialize(pWICBitmapSource, size.cx, size.cy, WICBitmapInterpolationModeFant);
			}
			if(SUCCEEDED(hr)) {
				WICRect rc = {0};
				if(needsFrame) {
					rc.Height = bitmapInfo.bmHeight;
					rc.Width = bitmapInfo.bmWidth;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					LPRGBQUAD pIconBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, rc.Width * rc.Height * sizeof(RGBQUAD)));
					hr = pWICBitmapSource->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pIconBits));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						// center the icon
						int xIconStart = (size.cx - bitmapInfo.bmWidth) / 2;
						int yIconStart = (size.cy - bitmapInfo.bmHeight) / 2;
						LPRGBQUAD pIconPixel = pIconBits;
						LPRGBQUAD pPixel = pBits;
						pPixel += yIconStart * size.cx;
						for(int y = yIconStart; y < yIconStart + bitmapInfo.bmHeight; ++y, pPixel += size.cx, pIconPixel += bitmapInfo.bmWidth) {
							CopyMemory(pPixel + xIconStart, pIconPixel, bitmapInfo.bmWidth * sizeof(RGBQUAD));
						}
						HeapFree(GetProcessHeap(), 0, pIconBits);

						rc.Height = size.cy;
						rc.Width = size.cx;

						// TODO: now draw a frame around it
					}
				} else {
					rc.Height = size.cy;
					rc.Width = size.cx;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					hr = pWICBitmapScaler->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pBits));
					ATLASSERT(SUCCEEDED(hr));

					if(SUCCEEDED(hr) && doAlphaChannelPostProcessing) {
						for(int i = 0; i < rc.Width * rc.Height; ++i, ++pBits) {
							if(pBits->rgbReserved == 0x00) {
								ZeroMemory(pBits, sizeof(RGBQUAD));
							}
						}
					}
				}
			} else {
				ATLASSERT(FALSE && "Bitmap scaler failed");
			}
		}
	}
	return hr;
}