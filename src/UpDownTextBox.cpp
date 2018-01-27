// UpDownTextBox.cpp: Combines Edit with msctls_updown32.

#include "stdafx.h"
#include "UpDownTextBox.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
IMEModeConstants UpDownTextBox::IMEFlags::chineseIMETable[10] = {
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

IMEModeConstants UpDownTextBox::IMEFlags::japaneseIMETable[10] = {
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

IMEModeConstants UpDownTextBox::IMEFlags::koreanIMETable[10] = {
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

FONTDESC UpDownTextBox::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


#pragma warning(push)
#pragma warning(disable: 4355)     // passing "this" to a member constructor
UpDownTextBox::UpDownTextBox() :
    containedEdit(WC_EDIT, this, 1),
    containedUpDown(UPDOWN_CLASS, this, 2)
{
	properties.font.InitializePropertyWatcher(this, DISPID_UPDWNTXTBOX_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_UPDWNTXTBOX_MOUSEICON);

	SIZEL size = {100, 20};
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
}
#pragma warning(pop)


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP UpDownTextBox::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IUpDownTextBox, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP UpDownTextBox::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<UpDownTextBox>::Load(pPropertyBag, pErrorLog);
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

STDMETHODIMP UpDownTextBox::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<UpDownTextBox>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
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

STDMETHODIMP UpDownTextBox::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 23 VT_I4 properties...
	pSize->QuadPart += 23 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 21 VT_BOOL properties...
	pSize->QuadPart += 21 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 2 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.cueBanner.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

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

	return S_OK;
}

STDMETHODIMP UpDownTextBox::Load(LPSTREAM pStream)
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
		if(version > 0x0102) {
			return E_FAIL;
		}
		LONG subSignature = 0;
		if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
			return hr;
		}
		if(subSignature != 0x04040404/*4x 0x03 (-> UpDownTextBox)*/) {
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
	VARIANT_BOOL alwaysShowSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL automaticallyCorrectValue = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL automaticallySetText = propertyValue.boolVal;
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
	BaseConstants base = static_cast<BaseConstants>(propertyValue.lVal);
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
	LONG currentValue = propertyValue.lVal;
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL groupDigits = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HAlignmentConstants hAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL hotTracking = propertyValue.boolVal;
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
	OLE_XSIZE_PIXELS leftMargin = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maximum = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maxTextLength = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG minimum = propertyValue.lVal;
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OrientationConstants orientation = static_cast<OrientationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processArrowKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL readOnlyTextBox = propertyValue.boolVal;
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	UpDownPositionConstants upDownPosition = static_cast<UpDownPositionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL wrapAtBoundaries = propertyValue.boolVal;

	VARIANT_BOOL detectDoubleClicks = VARIANT_FALSE;
	OLE_COLOR disabledBackColor = static_cast<OLE_COLOR>(-1);
	OLE_COLOR disabledForeColor = static_cast<OLE_COLOR>(-1);
	if(version >= 0x0101) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		disabledBackColor = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		disabledForeColor = propertyValue.lVal;

		if(version >= 0x0102) {
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
				return hr;
			}
			detectDoubleClicks = propertyValue.boolVal;
		}
	}


	hr = put_AcceptNumbersOnly(acceptNumbersOnly);
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
	hr = put_AutomaticallyCorrectValue(automaticallyCorrectValue);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutomaticallySetText(automaticallySetText);
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
	hr = put_Base(base);
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
	hr = put_CurrentValue(NULL, currentValue);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DetectDoubleClicks(detectDoubleClicks);
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
	hr = put_GroupDigits(groupDigits);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HAlignment(hAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotTracking(hotTracking);
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
	hr = put_LeftMargin(leftMargin);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Maximum(maximum);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxTextLength(maxTextLength);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Minimum(minimum);
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
	hr = put_Orientation(orientation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessArrowKeys(processArrowKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ReadOnlyTextBox(readOnlyTextBox);
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
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UpDownPosition(upDownPosition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_WrapAtBoundaries(wrapAtBoundaries);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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
	LONG version = 0x0102;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x04040404/*4x 0x03 (-> UpDownTextBox)*/;
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.automaticallyCorrectValue);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.automaticallySetText);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.autoScrolling;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.base;
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
	propertyValue.lVal = properties.currentValue;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
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
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.groupDigits);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.hotTracking);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.leftMargin;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maximum;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maxTextLength;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.minimum;
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
	propertyValue.lVal = properties.orientation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processArrowKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.readOnlyTextBox);
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
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
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
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.upDownPosition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.wrapAtBoundaries);
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
	// version 0x0102 starts here
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND UpDownTextBox::Create(HWND hWndParent, RECT& rect, LPARAM createParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_STANDARD_CLASSES | ICC_UPDOWN_CLASS;
	InitCommonControlsEx(&data);

	return CComCompositeControl<UpDownTextBox>::Create(hWndParent, rect, createParam);
}

HRESULT UpDownTextBox::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("UpDownTextBox ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void UpDownTextBox::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP UpDownTextBox::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComCompositeControl<UpDownTextBox>::IOleObject_SetClientSite(pClientSite);

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

	return hr;
}

STDMETHODIMP UpDownTextBox::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP UpDownTextBox::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
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
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
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

STDMETHODIMP UpDownTextBox::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
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
// implementation of ICategorizeProperties
STDMETHODIMP UpDownTextBox::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
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

STDMETHODIMP UpDownTextBox::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_UPDWNTXTBOX_ALWAYSSHOWSELECTION:
		case DISPID_UPDWNTXTBOX_APPEARANCE:
		case DISPID_UPDWNTXTBOX_BORDERSTYLE:
		case DISPID_UPDWNTXTBOX_CUEBANNER:
		case DISPID_UPDWNTXTBOX_DISPLAYCUEBANNERONFOCUS:
		case DISPID_UPDWNTXTBOX_FIRSTVISIBLECHAR:
		case DISPID_UPDWNTXTBOX_FORMATTINGRECTANGLEHEIGHT:
		case DISPID_UPDWNTXTBOX_FORMATTINGRECTANGLELEFT:
		case DISPID_UPDWNTXTBOX_FORMATTINGRECTANGLETOP:
		case DISPID_UPDWNTXTBOX_FORMATTINGRECTANGLEWIDTH:
		case DISPID_UPDWNTXTBOX_HALIGNMENT:
		case DISPID_UPDWNTXTBOX_LEFTMARGIN:
		case DISPID_UPDWNTXTBOX_MOUSEICON:
		case DISPID_UPDWNTXTBOX_MOUSEPOINTER:
		case DISPID_UPDWNTXTBOX_ORIENTATION:
		case DISPID_UPDWNTXTBOX_RIGHTMARGIN:
		case DISPID_UPDWNTXTBOX_SELECTEDTEXT:
		case DISPID_UPDWNTXTBOX_TEXT:
		case DISPID_UPDWNTXTBOX_UPDOWNPOSITION:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_UPDWNTXTBOX_ACCEPTNUMBERSONLY:
		case DISPID_UPDWNTXTBOX_AUTOMATICALLYCORRECTVALUE:
		case DISPID_UPDWNTXTBOX_AUTOMATICALLYSETTEXT:
		case DISPID_UPDWNTXTBOX_AUTOSCROLLING:
		case DISPID_UPDWNTXTBOX_CANCELIMECOMPOSITIONONSETFOCUS:
		case DISPID_UPDWNTXTBOX_CHARACTERCONVERSION:
		case DISPID_UPDWNTXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS:
		case DISPID_UPDWNTXTBOX_DETECTDOUBLECLICKS:
		case DISPID_UPDWNTXTBOX_DISABLEDEVENTS:
		case DISPID_UPDWNTXTBOX_DONTREDRAW:
		case DISPID_UPDWNTXTBOX_DOOEMCONVERSION:
		case DISPID_UPDWNTXTBOX_GROUPDIGITS:
		case DISPID_UPDWNTXTBOX_HOTTRACKING:
		case DISPID_UPDWNTXTBOX_HOVERTIME:
		case DISPID_UPDWNTXTBOX_IMEMODE:
		case DISPID_UPDWNTXTBOX_MAXTEXTLENGTH:
		case DISPID_UPDWNTXTBOX_PROCESSARROWKEYS:
		case DISPID_UPDWNTXTBOX_PROCESSCONTEXTMENUKEYS:
		case DISPID_UPDWNTXTBOX_RIGHTTOLEFT:
		case DISPID_UPDWNTXTBOX_WORDBREAKFUNCTION:
		case DISPID_UPDWNTXTBOX_WRAPATBOUNDARIES:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_UPDWNTXTBOX_BACKCOLOR:
		case DISPID_UPDWNTXTBOX_DISABLEDBACKCOLOR:
		case DISPID_UPDWNTXTBOX_DISABLEDFORECOLOR:
		case DISPID_UPDWNTXTBOX_FORECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_UPDWNTXTBOX_APPID:
		case DISPID_UPDWNTXTBOX_APPNAME:
		case DISPID_UPDWNTXTBOX_APPSHORTNAME:
		case DISPID_UPDWNTXTBOX_BUILD:
		case DISPID_UPDWNTXTBOX_CHARSET:
		case DISPID_UPDWNTXTBOX_HWND:
		case DISPID_UPDWNTXTBOX_HWNDEDIT:
		case DISPID_UPDWNTXTBOX_HWNDUPDOWN:
		case DISPID_UPDWNTXTBOX_ISRELEASE:
		case DISPID_UPDWNTXTBOX_MODIFIED:
		case DISPID_UPDWNTXTBOX_PROGRAMMER:
		case DISPID_UPDWNTXTBOX_TESTER:
		case DISPID_UPDWNTXTBOX_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_UPDWNTXTBOX_REGISTERFOROLEDRAGDROP:
		case DISPID_UPDWNTXTBOX_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_UPDWNTXTBOX_FONT:
		case DISPID_UPDWNTXTBOX_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_UPDWNTXTBOX_ENABLED:
		case DISPID_UPDWNTXTBOX_READONLYTEXTBOX:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
		case DISPID_UPDWNTXTBOX_ACCELERATORS:
		case DISPID_UPDWNTXTBOX_BASE:
		case DISPID_UPDWNTXTBOX_CURRENTVALUE:
		case DISPID_UPDWNTXTBOX_MAXIMUM:
		case DISPID_UPDWNTXTBOX_MINIMUM:
			*pCategory = PROPCAT_Scale;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString UpDownTextBox::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString UpDownTextBox::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString UpDownTextBox::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=EditControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString UpDownTextBox::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString UpDownTextBox::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString UpDownTextBox::GetVersion(void)
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
STDMETHODIMP UpDownTextBox::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_UPDWNTXTBOX_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_UPDWNTXTBOX_AUTOSCROLLING:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.autoScrolling), description);
			break;
		case DISPID_UPDWNTXTBOX_BASE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.base), description);
			break;
		case DISPID_UPDWNTXTBOX_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_UPDWNTXTBOX_CHARACTERCONVERSION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.characterConversion), description);
			break;
		case DISPID_UPDWNTXTBOX_CUEBANNER:
		case DISPID_UPDWNTXTBOX_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_UPDWNTXTBOX_HALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hAlignment), description);
			break;
		case DISPID_UPDWNTXTBOX_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_UPDWNTXTBOX_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_UPDWNTXTBOX_ORIENTATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.orientation), description);
			break;
		case DISPID_UPDWNTXTBOX_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_UPDWNTXTBOX_UPDOWNPOSITION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.upDownPosition), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<UpDownTextBox>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP UpDownTextBox::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_UPDWNTXTBOX_BASE:
		case DISPID_UPDWNTXTBOX_BORDERSTYLE:
		case DISPID_UPDWNTXTBOX_ORIENTATION:
		case DISPID_UPDWNTXTBOX_UPDOWNPOSITION:
			c = 2;
			break;
		case DISPID_UPDWNTXTBOX_APPEARANCE:
		case DISPID_UPDWNTXTBOX_CHARACTERCONVERSION:
		case DISPID_UPDWNTXTBOX_HALIGNMENT:
			c = 3;
			break;
		case DISPID_UPDWNTXTBOX_AUTOSCROLLING:
		case DISPID_UPDWNTXTBOX_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_UPDWNTXTBOX_IMEMODE:
			c = 12;
			break;
		case DISPID_UPDWNTXTBOX_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<UpDownTextBox>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = reinterpret_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = reinterpret_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_UPDWNTXTBOX_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if(property == DISPID_UPDWNTXTBOX_IMEMODE) {
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

STDMETHODIMP UpDownTextBox::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_UPDWNTXTBOX_APPEARANCE:
		case DISPID_UPDWNTXTBOX_AUTOSCROLLING:
		case DISPID_UPDWNTXTBOX_BASE:
		case DISPID_UPDWNTXTBOX_BORDERSTYLE:
		case DISPID_UPDWNTXTBOX_CHARACTERCONVERSION:
		case DISPID_UPDWNTXTBOX_HALIGNMENT:
		case DISPID_UPDWNTXTBOX_IMEMODE:
		case DISPID_UPDWNTXTBOX_MOUSEPOINTER:
		case DISPID_UPDWNTXTBOX_ORIENTATION:
		case DISPID_UPDWNTXTBOX_RIGHTTOLEFT:
		case DISPID_UPDWNTXTBOX_UPDOWNPOSITION:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<UpDownTextBox>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_UPDWNTXTBOX_CUEBANNER:
		case DISPID_UPDWNTXTBOX_TEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<UpDownTextBox>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT UpDownTextBox::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_UPDWNTXTBOX_APPEARANCE:
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
		case DISPID_UPDWNTXTBOX_AUTOSCROLLING:
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
		case DISPID_UPDWNTXTBOX_BASE:
			switch(cookie) {
				case bDecimal:
					description = GetResStringWithNumber(IDP_BASEDECIMAL, bDecimal);
					break;
				case bHexadecimal:
					description = GetResStringWithNumber(IDP_BASEHEXADECIMAL, bHexadecimal);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_UPDWNTXTBOX_BORDERSTYLE:
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
		case DISPID_UPDWNTXTBOX_CHARACTERCONVERSION:
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
		case DISPID_UPDWNTXTBOX_HALIGNMENT:
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
		case DISPID_UPDWNTXTBOX_IMEMODE:
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
		case DISPID_UPDWNTXTBOX_MOUSEPOINTER:
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
		case DISPID_UPDWNTXTBOX_ORIENTATION:
			switch(cookie) {
				case oHorizontal:
					description = GetResStringWithNumber(IDP_ORIENTATIONHORIZONTAL, oHorizontal);
					break;
				case oVertical:
					description = GetResStringWithNumber(IDP_ORIENTATIONVERTICAL, oVertical);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_UPDWNTXTBOX_RIGHTTOLEFT:
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
		case DISPID_UPDWNTXTBOX_UPDOWNPOSITION:
			switch(cookie) {
				case udLeftOfTextBox:
					description = GetResStringWithNumber(IDP_UPDOWNLEFTOFTEXTBOX, udLeftOfTextBox);
					break;
				case udRightOfTextBox:
					description = GetResStringWithNumber(IDP_UPDOWNRIGHTOFTEXTBOX, udRightOfTextBox);
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
STDMETHODIMP UpDownTextBox::GetPages(CAUUID* pPropertyPages)
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
STDMETHODIMP UpDownTextBox::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<UpDownTextBox>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP UpDownTextBox::EnumVerbs(IEnumOLEVERB** ppEnumerator)
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

HRESULT UpDownTextBox::DoVerbAbout(HWND hWndParent)
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

HRESULT UpDownTextBox::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_UPDWNTXTBOX_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT UpDownTextBox::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP UpDownTextBox::get_Accelerators(IUpDownAccelerators** ppAccelerators)
{
	ATLASSERT_POINTER(ppAccelerators, IUpDownAccelerators*);
	if(!ppAccelerators) {
		return E_POINTER;
	}

	ClassFactory::InitAccelerators(this, IID_IUpDownAccelerators, reinterpret_cast<LPUNKNOWN*>(ppAccelerators));
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AcceptNumbersOnly(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.acceptNumbersOnly = ((containedEdit.GetStyle() & ES_NUMBER) == ES_NUMBER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.acceptNumbersOnly);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_AcceptNumbersOnly(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_ACCEPTNUMBERSONLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.acceptNumbersOnly != b) {
		properties.acceptNumbersOnly = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			if(properties.acceptNumbersOnly) {
				containedEdit.ModifyStyle(0, ES_NUMBER);
			} else {
				containedEdit.ModifyStyle(ES_NUMBER, 0);
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_ACCEPTNUMBERSONLY);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AlwaysShowSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.alwaysShowSelection = ((containedEdit.GetStyle() & ES_NOHIDESEL) == ES_NOHIDESEL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_AlwaysShowSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_ALWAYSSHOWSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysShowSelection != b) {
		properties.alwaysShowSelection = b;
		SetDirty(TRUE);
		RecreateEditWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_ALWAYSSHOWSELECTION);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		if(containedEdit.GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(containedEdit.GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_APPEARANCE);
	if(newValue < 0 || newValue >= aDefault) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					containedEdit.ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					containedEdit.ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					containedEdit.ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}

			containedUpDown.DestroyWindow();
			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			containedEdit.MoveWindow(clientRectangle.left, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height());
			HWND hWndUD = CreateWindowEx(GetUpDownExStyleBits(), UPDOWN_CLASS, NULL, GetUpDownStyleBits(), clientRectangle.right, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height(), *this, NULL, ModuleHelper::GetModuleInstance(), NULL);
			containedUpDown.SubclassWindow(hWndUD);
			containedEdit.Invalidate();
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 12;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"EditControls");
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"EditCtls");
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AutomaticallyCorrectValue(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.automaticallyCorrectValue);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_AutomaticallyCorrectValue(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_AUTOMATICALLYCORRECTVALUE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.automaticallyCorrectValue != b) {
		properties.automaticallyCorrectValue = b;
		SetDirty(TRUE);

		if(properties.automaticallyCorrectValue && containedUpDown.IsWindow()) {
			/*if(IsComctl32Version580OrNewer()) {*/
				properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS32, 0, 0));
				containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
				if(((containedUpDown.GetStyle() & UDS_SETBUDDYINT) == UDS_SETBUDDYINT) && (GetFocus() == containedEdit)) {
					if(containedEdit.GetWindowTextLength() == 0) {
						/* HACK: If the contained edit control is empty and the value wasn't really changed, it's still
						         empty. */
						containedEdit.SetWindowText(TEXT("a"));
						containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
					}
				}
			//} else {
			//	properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS, 0, 0));
			//	containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<WORD>(properties.currentValue), 0));
			//	if(((containedUpDown.GetStyle() & UDS_SETBUDDYINT) == UDS_SETBUDDYINT) && (GetFocus() == containedEdit)) {
			//		if(containedEdit.GetWindowTextLength() == 0) {
			//			/* HACK: If the contained edit control is empty and the value wasn't really changed, it's still
			//			         empty. */
			//			containedEdit.SetWindowText(TEXT("a"));
			//			containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<WORD>(properties.currentValue), 0));
			//		}
			//	}
			//}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_AUTOMATICALLYCORRECTVALUE);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AutomaticallySetText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		properties.automaticallySetText = ((containedUpDown.GetStyle() & UDS_SETBUDDYINT) == UDS_SETBUDDYINT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.automaticallySetText);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_AutomaticallySetText(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_AUTOMATICALLYSETTEXT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.automaticallySetText != b) {
		properties.automaticallySetText = b;
		SetDirty(TRUE);
		RecreateUpDownWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_AUTOMATICALLYSETTEXT);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_AutoScrolling(AutoScrollingConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AutoScrollingConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
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

STDMETHODIMP UpDownTextBox::put_AutoScrolling(AutoScrollingConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_AUTOSCROLLING);
	if(properties.autoScrolling != newValue) {
		properties.autoScrolling = newValue;
		SetDirty(TRUE);
		RecreateEditWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_AUTOSCROLLING);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			properties.enabled = !(GetStyle() & WS_DISABLED);
		}
		if(!containedEdit.IsWindow()) {
			properties.readOnlyTextBox = ((containedEdit.GetStyle() & ES_READONLY) == ES_READONLY);
		}
		if(properties.enabled && !properties.readOnlyTextBox) {
			if(containedEdit.IsWindow()) {
				containedEdit.Invalidate();
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Base(BaseConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BaseConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		switch(containedUpDown.SendMessage(UDM_GETBASE, 0, 0)) {
			case 10:
				properties.base = bDecimal;
				break;
			case 16:
				properties.base = bHexadecimal;
				break;
			default:
				ATLASSERT(FALSE);
				break;
		}
	}

	*pValue = properties.base;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_Base(BaseConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_BASE);
	if(properties.base != newValue) {
		properties.base = newValue;
		SetDirty(TRUE);

		if(containedUpDown.IsWindow()) {
			switch(properties.base) {
				case bDecimal:
					containedUpDown.SendMessage(UDM_SETBASE, 10, 0);
					break;
				case bHexadecimal:
					containedUpDown.SendMessage(UDM_SETBASE, 16, 0);
					break;
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_BASE);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.borderStyle = ((containedEdit.GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					containedEdit.ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					containedEdit.ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}

			containedUpDown.DestroyWindow();
			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			containedEdit.MoveWindow(clientRectangle.left, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height());
			HWND hWndUD = CreateWindowEx(GetUpDownExStyleBits(), UPDOWN_CLASS, NULL, GetUpDownStyleBits(), clientRectangle.right, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height(), *this, NULL, ModuleHelper::GetModuleInstance(), NULL);
			containedUpDown.SubclassWindow(hWndUD);
			containedEdit.Invalidate();
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_CancelIMECompositionOnSetFocus(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.cancelIMECompositionOnSetFocus = ((containedEdit.SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0) & EIMES_CANCELCOMPSTRINFOCUS) == EIMES_CANCELCOMPSTRINFOCUS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.cancelIMECompositionOnSetFocus);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_CancelIMECompositionOnSetFocus(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_CANCELIMECOMPOSITIONONSETFOCUS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.cancelIMECompositionOnSetFocus != b) {
		properties.cancelIMECompositionOnSetFocus = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			DWORD flags = static_cast<DWORD>(containedEdit.SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0));
			if(properties.cancelIMECompositionOnSetFocus) {
				flags |= EIMES_CANCELCOMPSTRINFOCUS;
			} else {
				flags &= ~EIMES_CANCELCOMPSTRINFOCUS;
			}
			containedEdit.SendMessage(EM_SETIMESTATUS, EMSIS_COMPOSITIONSTRING, flags);
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_CANCELIMECOMPOSITIONONSETFOCUS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_CharacterConversion(CharacterConversionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, CharacterConversionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		DWORD style = containedEdit.GetStyle();
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

STDMETHODIMP UpDownTextBox::put_CharacterConversion(CharacterConversionConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_CHARACTERCONVERSION);
	if(properties.characterConversion != newValue) {
		properties.characterConversion = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			switch(properties.characterConversion) {
				case ccNone:
					containedEdit.ModifyStyle(ES_LOWERCASE | ES_UPPERCASE, 0);
					break;
				case ccLowerCase:
					containedEdit.ModifyStyle(ES_UPPERCASE, ES_LOWERCASE);
					break;
				case ccUpperCase:
					containedEdit.ModifyStyle(ES_LOWERCASE, ES_UPPERCASE);
					break;
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_CHARACTERCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_CharSet(BSTR* pValue)
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

STDMETHODIMP UpDownTextBox::get_CompleteIMECompositionOnKillFocus(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.completeIMECompositionOnKillFocus = ((containedEdit.SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0) & EIMES_COMPLETECOMPSTRKILLFOCUS) == EIMES_COMPLETECOMPSTRKILLFOCUS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.completeIMECompositionOnKillFocus);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_CompleteIMECompositionOnKillFocus(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.completeIMECompositionOnKillFocus != b) {
		properties.completeIMECompositionOnKillFocus = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			DWORD flags = static_cast<DWORD>(containedEdit.SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0));
			if(properties.completeIMECompositionOnKillFocus) {
				flags |= EIMES_COMPLETECOMPSTRKILLFOCUS;
			} else {
				flags &= ~EIMES_COMPLETECOMPSTRKILLFOCUS;
			}
			containedEdit.SendMessage(EM_SETIMESTATUS, EMSIS_COMPOSITIONSTRING, flags);
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_CueBanner(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		WCHAR pBuffer[1025];
		if(containedEdit.SendMessage(EM_GETCUEBANNER, reinterpret_cast<WPARAM>(pBuffer), 1024)) {
			properties.cueBanner = pBuffer;
		}
	}

	*pValue = properties.cueBanner.Copy();
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_CueBanner(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_CUEBANNER);
	if(properties.cueBanner != newValue) {
		properties.cueBanner.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(containedEdit.IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			containedEdit.SendMessage(EM_SETCUEBANNER, properties.displayCueBannerOnFocus, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_CUEBANNER);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_CurrentValue(VARIANT_BOOL* pInvalidValue/* = NULL*/, LONG* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	BOOL error = FALSE;
	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		/*if(IsComctl32Version580OrNewer()) {*/
			properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS32, 0, reinterpret_cast<LPARAM>(&error)));
		/*} else {
			properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS, 0, 0));
		}*/
	}

	*pValue = properties.currentValue;
	if(pInvalidValue) {
		*pInvalidValue = BOOL2VARIANTBOOL(error);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_CurrentValue(VARIANT_BOOL* /*pInvalidValue = NULL*/, LONG newValue/* = 0*/)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_CURRENTVALUE);
	if(properties.currentValue != newValue) {
		properties.currentValue = newValue;
		SetDirty(TRUE);

		if(containedUpDown.IsWindow()) {
			/*if(IsComctl32Version580OrNewer()) {*/
				containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
			/*} else {
				containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<SHORT>(properties.currentValue), 0));
			}*/
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_CURRENTVALUE);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_DetectDoubleClicks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_DetectDoubleClicks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_DETECTDOUBLECLICKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDoubleClicks != b) {
		properties.detectDoubleClicks = b;
		SetDirty(TRUE);

		FireOnChanged(DISPID_UPDWNTXTBOX_DETECTDOUBLECLICKS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_DisabledBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledBackColor;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_DisabledBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_DISABLEDBACKCOLOR);
	if(properties.disabledBackColor != newValue) {
		properties.disabledBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			properties.enabled = !(GetStyle() & WS_DISABLED);
		}
		if(!containedEdit.IsWindow()) {
			properties.readOnlyTextBox = ((containedEdit.GetStyle() & ES_READONLY) == ES_READONLY);
		}
		if(!properties.enabled || properties.readOnlyTextBox) {
			if(IsWindow()) {
				CRect windowRectangle;
				containedEdit.GetWindowRect(&windowRectangle);
				WINDOWPOS details = {0};
				details.hwnd = containedEdit;
				details.cx = windowRectangle.Width();
				details.cy = windowRectangle.Height();
				details.flags = SWP_DRAWFRAME | SWP_FRAMECHANGED | SWP_NOCOPYBITS | SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER;
				containedEdit.DefWindowProc(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&details));
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_DISABLEDBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((static_cast<long>(properties.disabledEvents) & deMouseEvents) != (static_cast<long>(newValue) & deMouseEvents)) {
			if(IsWindow()) {
				if(static_cast<long>(newValue) & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					trackingOptions.hwndTrack = containedEdit;
					TrackMouseEvent(&trackingOptions);
					trackingOptions.hwndTrack = containedUpDown;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		SetDirty(TRUE);
		properties.disabledEvents = newValue;
		FireOnChanged(DISPID_UPDWNTXTBOX_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_DisabledForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledForeColor;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_DisabledForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_DISABLEDFORECOLOR);
	if(properties.disabledForeColor != newValue) {
		properties.disabledForeColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_UPDWNTXTBOX_DISABLEDFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_DisplayCueBannerOnFocus(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayCueBannerOnFocus);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_DisplayCueBannerOnFocus(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_DISPLAYCUEBANNERONFOCUS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayCueBannerOnFocus != b) {
		properties.displayCueBannerOnFocus = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			containedEdit.SendMessage(EM_SETCUEBANNER, properties.displayCueBannerOnFocus, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_DISPLAYCUEBANNERONFOCUS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		if(containedEdit.IsWindow()) {
			containedEdit.SetRedraw(!b);
		}
		if(containedUpDown.IsWindow()) {
			containedUpDown.SetRedraw(!b);
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_DoOEMConversion(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.doOEMConversion = ((containedEdit.GetStyle() & ES_OEMCONVERT) == ES_OEMCONVERT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.doOEMConversion);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_DoOEMConversion(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_DOOEMCONVERSION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.doOEMConversion != b) {
		properties.doOEMConversion = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			if(properties.doOEMConversion) {
				containedEdit.ModifyStyle(0, ES_OEMCONVERT);
			} else {
				containedEdit.ModifyStyle(ES_OEMCONVERT, 0);
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_DOOEMCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Enabled(VARIANT_BOOL* pValue)
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

STDMETHODIMP UpDownTextBox::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(!containedEdit.IsWindow()) {
			properties.readOnlyTextBox = ((containedEdit.GetStyle() & ES_READONLY) == ES_READONLY);
		}
		if(IsWindow()) {
			EnableWindow(properties.enabled);
		}
		if(containedEdit.IsWindow()) {
			containedEdit.EnableWindow(properties.enabled);
		}
		if(containedUpDown.IsWindow()) {
			containedUpDown.EnableWindow(properties.enabled);
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_UPDWNTXTBOX_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_FirstVisibleChar(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		if((containedEdit.GetStyle() & ES_MULTILINE) == 0) {
			*pValue = static_cast<LONG>(containedEdit.SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0));
		}
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Font(IFontDisp** ppFont)
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

STDMETHODIMP UpDownTextBox::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_FONT);
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
	FireOnChanged(DISPID_UPDWNTXTBOX_FONT);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_FONT);
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
	FireOnChanged(DISPID_UPDWNTXTBOX_FONT);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.Invalidate();
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_FormattingRectangleHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		CRect formattingRectangle;
		containedEdit.SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formattingRectangle));
		*pValue = formattingRectangle.Height();
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_FormattingRectangleLeft(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		CRect formattingRectangle;
		containedEdit.SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formattingRectangle));
		*pValue = formattingRectangle.left;
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_FormattingRectangleTop(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		CRect formattingRectangle;
		containedEdit.SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formattingRectangle));
		*pValue = formattingRectangle.top;
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_FormattingRectangleWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		CRect formattingRectangle;
		containedEdit.SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formattingRectangle));
		*pValue = formattingRectangle.Width();
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_GroupDigits(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		properties.groupDigits = ((containedUpDown.GetStyle() & UDS_NOTHOUSANDS) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.groupDigits);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_GroupDigits(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_GROUPDIGITS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.groupDigits != b) {
		properties.groupDigits = b;
		SetDirty(TRUE);
		RecreateUpDownWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_GROUPDIGITS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		switch(containedEdit.GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
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

STDMETHODIMP UpDownTextBox::put_HAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_HALIGNMENT);
	if(properties.hAlignment != newValue) {
		properties.hAlignment = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			if(RunTimeHelper::IsVista()) {
				switch(properties.hAlignment) {
					case halLeft:
						containedEdit.ModifyStyle(ES_CENTER | ES_RIGHT, ES_LEFT);
						break;
					case halCenter:
						containedEdit.ModifyStyle(ES_LEFT | ES_RIGHT, ES_CENTER);
						break;
					case halRight:
						containedEdit.ModifyStyle(ES_LEFT | ES_CENTER, ES_RIGHT);
						break;
				}
				containedEdit.Invalidate();
			} else {
				RecreateEditWindow();
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_HALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_HotTracking(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		properties.hotTracking = ((containedUpDown.GetStyle() & UDS_HOTTRACK) == UDS_HOTTRACK);
	}

	*pValue = BOOL2VARIANTBOOL(properties.hotTracking);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_HotTracking(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_HOTTRACKING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.hotTracking != b) {
		properties.hotTracking = b;
		SetDirty(TRUE);
		RecreateUpDownWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_HOTTRACKING);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_UPDWNTXTBOX_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_hWndEdit(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(containedEdit));
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_hWndUpDown(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(containedUpDown));
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_IMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.IMEMode;
	} else {
		if((GetFocus() == containedEdit) && (GetEffectiveIMEMode() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(containedEdit);
			if((ime != imeInherit) && (properties.IMEMode != imeInherit)) {
				properties.IMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode();
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == containedEdit) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(containedEdit, ime);
				}
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_IsRelease(VARIANT_BOOL* pValue)
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

STDMETHODIMP UpDownTextBox::get_LeftMargin(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.leftMargin = LOWORD(containedEdit.SendMessage(EM_GETMARGINS, 0, 0));
	}

	*pValue = properties.leftMargin;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_LeftMargin(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_LEFTMARGIN);
	if(properties.leftMargin != newValue) {
		properties.leftMargin = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.SendMessage(EM_SETMARGINS, EC_LEFTMARGIN, MAKELPARAM((properties.leftMargin == -1 ? EC_USEFONTINFO : properties.leftMargin), 0));
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_LEFTMARGIN);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Maximum(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		containedUpDown.SendMessage(UDM_GETRANGE32, 0, reinterpret_cast<LPARAM>(&properties.maximum));
	}

	*pValue = properties.maximum;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_Maximum(LONG newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_MAXIMUM);
	if(properties.maximum != newValue) {
		properties.maximum = newValue;
		SetDirty(TRUE);

		if(containedUpDown.IsWindow()) {
			containedUpDown.SendMessage(UDM_GETRANGE32, reinterpret_cast<WPARAM>(&properties.minimum), 0);
			containedUpDown.SendMessage(UDM_SETRANGE32, properties.minimum, properties.maximum);
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_MAXIMUM);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_MaxTextLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.maxTextLength = static_cast<LONG>(containedEdit.SendMessage(EM_GETLIMITTEXT, 0, 0));
	}

	*pValue = properties.maxTextLength;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_MaxTextLength(LONG newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_MAXTEXTLENGTH);
	if(properties.maxTextLength != newValue) {
		properties.maxTextLength = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.SendMessage(EM_SETLIMITTEXT, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength), 0);
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_MAXTEXTLENGTH);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Minimum(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		containedUpDown.SendMessage(UDM_GETRANGE32, reinterpret_cast<WPARAM>(&properties.minimum), 0);
	}

	*pValue = properties.minimum;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_Minimum(LONG newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_MINIMUM);
	if(properties.minimum != newValue) {
		properties.minimum = newValue;
		SetDirty(TRUE);

		if(containedUpDown.IsWindow()) {
			containedUpDown.SendMessage(UDM_GETRANGE32, 0, reinterpret_cast<LPARAM>(&properties.maximum));
			containedUpDown.SendMessage(UDM_SETRANGE32, properties.minimum, properties.maximum);
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_MINIMUM);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Modified(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.modified = containedEdit.SendMessage(EM_GETMODIFY, 0, 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.modified);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_Modified(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_MODIFIED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.modified != b) {
		properties.modified = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.SendMessage(EM_SETMODIFY, properties.modified, 0);
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_MODIFIED);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_MouseIcon(IPictureDisp** ppMouseIcon)
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

STDMETHODIMP UpDownTextBox::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_MOUSEICON);
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
	FireOnChanged(DISPID_UPDWNTXTBOX_MOUSEICON);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_MOUSEICON);
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
	FireOnChanged(DISPID_UPDWNTXTBOX_MOUSEICON);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_UPDWNTXTBOX_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Orientation(OrientationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OrientationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		properties.orientation = ((containedUpDown.GetStyle() & UDS_HORZ) == UDS_HORZ ? oHorizontal : oVertical);
	}

	*pValue = properties.orientation;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_Orientation(OrientationConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_ORIENTATION);
	if(properties.orientation != newValue) {
		properties.orientation = newValue;
		SetDirty(TRUE);
		RecreateUpDownWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_ORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_ProcessArrowKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		properties.processArrowKeys = ((containedUpDown.GetStyle() & UDS_ARROWKEYS) == UDS_ARROWKEYS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.processArrowKeys);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_ProcessArrowKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_PROCESSARROWKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processArrowKeys != b) {
		properties.processArrowKeys = b;
		SetDirty(TRUE);
		RecreateUpDownWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_PROCESSARROWKEYS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_UPDWNTXTBOX_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_ReadOnlyTextBox(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.readOnlyTextBox = ((containedEdit.GetStyle() & ES_READONLY) == ES_READONLY);
	}

	*pValue = BOOL2VARIANTBOOL(properties.readOnlyTextBox);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_ReadOnlyTextBox(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_READONLYTEXTBOX);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.readOnlyTextBox != b) {
		properties.readOnlyTextBox = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(containedEdit.IsWindow()) {
				containedEdit.SendMessage(EM_SETREADONLY, properties.readOnlyTextBox, 0);
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_READONLYTEXTBOX);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_RightMargin(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		properties.rightMargin = HIWORD(containedEdit.SendMessage(EM_GETMARGINS, 0, 0));
	}

	*pValue = properties.rightMargin;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_RightMargin(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_RIGHTMARGIN);
	if(properties.rightMargin != newValue) {
		properties.rightMargin = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.SendMessage(EM_SETMARGINS, EC_RIGHTMARGIN, MAKELPARAM(0, (properties.rightMargin == -1 ? EC_USEFONTINFO : properties.rightMargin)));
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_RIGHTMARGIN);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(containedEdit, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(containedEdit) == S_OK);
			}
		}
		if(containedUpDown.IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(containedUpDown, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(containedUpDown) == S_OK);
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		// TODO: Process WM_STYLECHANGED to automatically change the up down controls' styles.
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = containedEdit.GetExStyle();
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

STDMETHODIMP UpDownTextBox::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow() && containedEdit.IsWindow() && containedUpDown.IsWindow()) {
			if(properties.rightToLeft & rtlLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL);
				containedEdit.ModifyStyleEx(0, WS_EX_LAYOUTRTL);
				containedUpDown.ModifyStyleEx(0, WS_EX_LAYOUTRTL);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
				containedEdit.ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
				containedUpDown.ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
			}
			if(properties.rightToLeft & rtlText) {
				ModifyStyleEx(0, WS_EX_RTLREADING);
				containedEdit.ModifyStyleEx(0, WS_EX_RTLREADING);
				containedUpDown.ModifyStyleEx(0, WS_EX_RTLREADING);
			} else {
				ModifyStyleEx(WS_EX_RTLREADING, 0);
				containedEdit.ModifyStyleEx(WS_EX_RTLREADING, 0);
				containedUpDown.ModifyStyleEx(WS_EX_RTLREADING, 0);
			}

			switch(containedEdit.GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
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
			// size the edit control to 100% of the container
			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			containedEdit.MoveWindow(clientRectangle.left, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height());
			// now reset the buddy, so that the up down control places itself into the edit control
			containedUpDown.SendMessage(UDM_SETBUDDY, reinterpret_cast<WPARAM>(static_cast<HWND>(containedEdit)), 0);

			/* TODO: Why is this necessary? If we're using RTL features, we sometimes end up with an edit control
			         that has this style set. Why? Alignment is changed, too. */
			containedEdit.ModifyStyleEx(WS_EX_RIGHT, 0);
			switch(properties.hAlignment) {
				case halLeft:
					containedEdit.ModifyStyle(ES_CENTER | ES_RIGHT, ES_LEFT);
					break;
				case halCenter:
					containedEdit.ModifyStyle(ES_LEFT | ES_RIGHT, ES_CENTER);
					break;
				case halRight:
					containedEdit.ModifyStyle(ES_LEFT | ES_CENTER, ES_RIGHT);
					break;
			}
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_SelectedText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		int bufferSize = containedEdit.GetWindowTextLength() + 1;
		LPTSTR pBuffer = new TCHAR[bufferSize];
		containedEdit.GetWindowText(pBuffer, bufferSize);

		int selectionStart = 0;
		int selectionEnd = 0;
		containedEdit.SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
		pBuffer[selectionEnd] = TEXT('\0');
		*pValue = _bstr_t(&pBuffer[selectionStart]).Detach();

		delete[] pBuffer;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_UPDWNTXTBOX_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedEdit.IsWindow()) {
		containedEdit.GetWindowText(&properties.text);
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_TEXT);
	properties.text.AssignBSTR(newValue);
	if(containedEdit.IsWindow()) {
		SetDirty(TRUE);
		containedEdit.SetWindowText(COLE2CT(properties.text));
		FireOnChanged(DISPID_UPDWNTXTBOX_TEXT);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_UpDownPosition(UpDownPositionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UpDownPositionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		switch(containedUpDown.GetStyle() & (UDS_ALIGNLEFT | UDS_ALIGNRIGHT)) {
			case UDS_ALIGNLEFT:
				properties.upDownPosition = udLeftOfTextBox;
				break;
			case UDS_ALIGNRIGHT:
				properties.upDownPosition = udRightOfTextBox;
				break;
		}
	}

	*pValue = properties.upDownPosition;
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_UpDownPosition(UpDownPositionConstants newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_UPDOWNPOSITION);
	if(properties.upDownPosition != newValue) {
		properties.upDownPosition = newValue;
		SetDirty(TRUE);
		RecreateUpDownWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_UPDOWNPOSITION);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_UPDWNTXTBOX_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_Version(BSTR* pValue)
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

STDMETHODIMP UpDownTextBox::get_WordBreakFunction(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		*pValue = static_cast<LONG>(containedEdit.SendMessage(EM_GETWORDBREAKPROC, 0, 0));
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_WordBreakFunction(LONG newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_WORDBREAKFUNCTION);
	if(containedEdit.IsWindow()) {
		containedEdit.SendMessage(EM_SETWORDBREAKPROC, 0, static_cast<LPARAM>(newValue));
	}

	FireOnChanged(DISPID_UPDWNTXTBOX_WORDBREAKFUNCTION);
	FireViewChange();
	SendOnDataChange(NULL);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::get_WrapAtBoundaries(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedUpDown.IsWindow()) {
		properties.wrapAtBoundaries = ((containedUpDown.GetStyle() & UDS_WRAP) == UDS_WRAP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.wrapAtBoundaries);
	return S_OK;
}

STDMETHODIMP UpDownTextBox::put_WrapAtBoundaries(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_UPDWNTXTBOX_WRAPATBOUNDARIES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.wrapAtBoundaries != b) {
		properties.wrapAtBoundaries = b;
		SetDirty(TRUE);
		RecreateUpDownWindow();
		FireOnChanged(DISPID_UPDWNTXTBOX_WRAPATBOUNDARIES);
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP UpDownTextBox::CanUndo(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		*pValue = BOOL2VARIANTBOOL(containedEdit.SendMessage(EM_CANUNDO, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::CharIndexToPosition(LONG characterIndex, OLE_XPOS_PIXELS* pX/* = NULL*/, OLE_YPOS_PIXELS* pY/* = NULL*/)
{
	if(containedEdit.IsWindow()) {
		LRESULT lr = containedEdit.SendMessage(EM_POSFROMCHAR, characterIndex, 0);
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

STDMETHODIMP UpDownTextBox::EmptyUndoBuffer(void)
{
	if(containedEdit.IsWindow()) {
		containedEdit.SendMessage(EM_EMPTYUNDOBUFFER, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::FinishOLEDragDrop(void)
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

STDMETHODIMP UpDownTextBox::GetSelection(LONG* pSelectionStart/* = NULL*/, LONG* pSelectionEnd/* = NULL*/)
{
	if(containedEdit.IsWindow()) {
		containedEdit.SendMessage(EM_GETSEL, reinterpret_cast<WPARAM>(pSelectionStart), reinterpret_cast<LPARAM>(pSelectionEnd));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::HideBalloonTip(VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		*pSucceeded = BOOL2VARIANTBOOL(containedEdit.SendMessage(EM_HIDEBALLOONTIP, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails)
{
	ATLASSERT_POINTER(pHitTestDetails, HitTestConstants);
	if(!pHitTestDetails) {
		return E_POINTER;
	}

	if(IsWindow()) {
		POINT pt = {x, y};
		*pHitTestDetails = htNotOverControl;

		CRect rc;
		GetWindowRect(&rc);
		ScreenToClient(&rc);
		if(rc.PtInRect(pt)) {
			containedEdit.GetWindowRect(&rc);
			ScreenToClient(&rc);
			if(rc.PtInRect(pt)) {
				*pHitTestDetails = htTextBox;
			} else {
				containedUpDown.GetWindowRect(&rc);
				ScreenToClient(&rc);
				if(rc.PtInRect(pt)) {
					*pHitTestDetails = htUpDown;
				}
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::IsTextTruncated(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		*pValue = VARIANT_FALSE;

		int bufferSize = containedEdit.GetWindowTextLength();
		if(bufferSize > 0) {
			LPTSTR pBuffer = new TCHAR[bufferSize + 1];
			containedEdit.GetWindowText(pBuffer, bufferSize + 1);

			// calculate the width
			CRect clientRectangle;
			containedEdit.GetClientRect(&clientRectangle);
			DWORD margins = static_cast<DWORD>(containedEdit.SendMessage(EM_GETMARGINS, 0, 0));
			clientRectangle.left += LOWORD(margins);
			clientRectangle.right -= HIWORD(margins);

			int firstCharacter = 0;
			int lastCharacter = firstCharacter + bufferSize - 1;
			CRect textRectangle;
			textRectangle.left = GET_X_LPARAM(containedEdit.SendMessage(EM_POSFROMCHAR, firstCharacter, 0));
			textRectangle.right = GET_X_LPARAM(containedEdit.SendMessage(EM_POSFROMCHAR, lastCharacter, 0));

			// add the last character's width
			CDCHandle targetDC = containedEdit.GetDC();
			HFONT hFont = reinterpret_cast<HFONT>(containedEdit.SendMessage(WM_GETFONT, 0, 0));
			HFONT hPreviousFont = targetDC.SelectFont(hFont);

			CRect rc = clientRectangle;
			switch(containedEdit.GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
				case ES_CENTER:
					targetDC.DrawText(&pBuffer[bufferSize - 1], 1, &rc, DT_CALCRECT | DT_CENTER | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
					break;
				case ES_RIGHT:
					targetDC.DrawText(&pBuffer[bufferSize - 1], 1, &rc, DT_CALCRECT | DT_RIGHT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
					break;
				case ES_LEFT:
					targetDC.DrawText(&pBuffer[bufferSize - 1], 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_EDITCONTROL | DT_SINGLELINE | DT_HIDEPREFIX);
					break;
			}
			textRectangle.right += rc.Width();

			targetDC.SelectFont(hPreviousFont);
			containedEdit.ReleaseDC(targetDC);

			*pValue = BOOL2VARIANTBOOL((textRectangle.left < clientRectangle.left) || (textRectangle.right > clientRectangle.right));
			delete[] pBuffer;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP UpDownTextBox::PositionToCharIndex(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG* pCharacterIndex/* = NULL*/)
{
	if(containedEdit.IsWindow()) {
		if(pCharacterIndex) {
			*pCharacterIndex = -1;
		}
		LRESULT lr = containedEdit.SendMessage(EM_CHARFROMPOS, 0, MAKELPARAM(x, y));
		if((LOWORD(lr) != 65535) || (HIWORD(lr) != 65535)) {
			if(pCharacterIndex) {
				*pCharacterIndex = LOWORD(lr);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::Refresh(void)
{
	if(containedEdit.IsWindow()) {
		containedEdit.Invalidate();
		containedEdit.UpdateWindow();
	}
	if(containedUpDown.IsWindow()) {
		containedUpDown.Invalidate();
		containedUpDown.UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP UpDownTextBox::ReplaceSelectedText(BSTR replacementText, VARIANT_BOOL undoable/* = VARIANT_FALSE*/)
{
	if(containedEdit.IsWindow()) {
		containedEdit.SendMessage(EM_REPLACESEL, VARIANTBOOL2BOOL(undoable), reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(COLE2CT(replacementText))));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP UpDownTextBox::ScrollCaretIntoView(void)
{
	if(containedEdit.IsWindow()) {
		containedEdit.SendMessage(EM_SCROLLCARET, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::SetSelection(LONG selectionStart, LONG selectionEnd)
{
	if(containedEdit.IsWindow()) {
		containedEdit.SendMessage(EM_SETSEL, selectionStart, selectionEnd);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::ShowBalloonTip(BSTR title, BSTR text, BalloonTipIconConstants icon/* = btiNone*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		EDITBALLOONTIP balloonTipDetails = {0};
		balloonTipDetails.cbStruct = sizeof(EDITBALLOONTIP);
		balloonTipDetails.pszText = OLE2W(text);
		balloonTipDetails.pszTitle = OLE2W(title);
		balloonTipDetails.ttiIcon = icon;
		*pSucceeded = BOOL2VARIANTBOOL(containedEdit.SendMessage(EM_SHOWBALLOONTIP, 0, reinterpret_cast<LPARAM>(&balloonTipDetails)));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP UpDownTextBox::Undo(VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	if(containedEdit.IsWindow()) {
		*pSucceeded = BOOL2VARIANTBOOL(containedEdit.SendMessage(EM_UNDO, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}


LRESULT UpDownTextBox::OnCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
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

LRESULT UpDownTextBox::OnCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(properties.disabledForeColor != static_cast<OLE_COLOR>(-1)) {
		SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.disabledForeColor));
	}

	if(RunTimeHelper::IsCommCtrl6() && properties.disabledBackColor == static_cast<OLE_COLOR>(-1)) {
		BOOL useTransparentTextBackground = FALSE;
		if(flags.usingThemes) {
			RECT clientRectangle;
			::GetClientRect(reinterpret_cast<HWND>(lParam), &clientRectangle);
			int state = SaveDC(reinterpret_cast<HDC>(wParam));				// if we don't do this, we get a large font if the control is read-only and sits inside a VB6 frame
			if(RunTimeHelper::IsVista()) {
				// NOTE: This solution produces more flickering then using a pattern brush.
				useTransparentTextBackground = (DrawThemeParentBackgroundEx(reinterpret_cast<HWND>(lParam), reinterpret_cast<HDC>(wParam), 0, &clientRectangle) == S_OK);
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
					useTransparentTextBackground = (DrawThemeParentBackground(reinterpret_cast<HWND>(lParam), memoryDC, &clientRectangle) == S_OK);

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
			useTransparentTextBackground = FALSE;
		}
		if(useTransparentTextBackground) {
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

LRESULT UpDownTextBox::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT UpDownTextBox::OnInitDialog(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// this will keep the object alive if the client destroys the control window in an event handler
	AddRef();

	if(properties.rightToLeft & rtlLayout) {
		ModifyStyleEx(0, WS_EX_LAYOUTRTL);
	} else {
		ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
	}

	CRect clientRectangle;
	GetClientRect(&clientRectangle);
	HWND hWndEdit = CreateWindowEx(GetEditExStyleBits(), WC_EDIT, NULL, GetEditStyleBits(), clientRectangle.left, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height(), *this, NULL, ModuleHelper::GetModuleInstance(), NULL);
	containedEdit.SubclassWindow(hWndEdit);
	HWND hWndUD = CreateWindowEx(GetUpDownExStyleBits(), UPDOWN_CLASS, NULL, GetUpDownStyleBits(), clientRectangle.right, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height(), *this, NULL, ModuleHelper::GetModuleInstance(), NULL);
	containedUpDown.SubclassWindow(hWndUD);

	/* TODO: Why is this necessary? If we're using RTL features, we sometimes end up with an edit control
	         that has this style set. Why? */
	containedEdit.ModifyStyleEx(WS_EX_RIGHT, 0);
	containedEdit.ModifyStyle(ES_RIGHT, 0);

	// we want to receive WM_*BUTTONDBLCLK messages
	// NOTE: We're changing the class style of UPDOWN_CLASS! Fortunately this change is app-wide only.
	DWORD style = GetClassLong(hWndUD, GCL_STYLE);
	style |= CS_DBLCLKS;
	SetClassLong(hWndUD, GCL_STYLE, style);

	Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	return FALSE;
}

LRESULT UpDownTextBox::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT UpDownTextBox::OnScroll(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deValueChangingEvents)) {
		Raise_ValueChanged();
	}
	return 0;
}

LRESULT UpDownTextBox::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
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
		HWND hWindowBelowMouse = WindowFromPoint(mousePosition);
		if((hWindowBelowMouse == *this) || IsChild(hWindowBelowMouse)) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
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

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT UpDownTextBox::OnDeltaPosNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	if(!(properties.disabledEvents & deValueChangingEvents)) {
		if(FireOnRequestEdit(DISPID_UPDWNTXTBOX_CURRENTVALUE) == S_FALSE) {
			return TRUE;
		}

		LPNMUPDOWN pDetails = reinterpret_cast<LPNMUPDOWN>(pNotificationDetails);
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_ValueChanging(pDetails->iPos, pDetails->iDelta, &cancel);
		if(cancel != VARIANT_FALSE) {
			return TRUE;
		}

		SetDirty(TRUE);
		FireOnChanged(DISPID_UPDWNTXTBOX_CURRENTVALUE);
		SendOnDataChange();
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT UpDownTextBox::OnAlign(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
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

LRESULT UpDownTextBox::OnChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(!(properties.disabledEvents & deTextChangedEvents)) {
		Raise_TextChanged();
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_UPDWNTXTBOX_TEXT);
	SendOnDataChange();

	wasHandled = FALSE;
	return 0;
}

LRESULT UpDownTextBox::OnErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_OutOfMemory();
	return 0;
}

LRESULT UpDownTextBox::OnKillFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(properties.automaticallyCorrectValue) {
		/*if(IsComctl32Version580OrNewer()) {*/
			properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS32, 0, 0));
			containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
			if(containedUpDown.GetStyle() & UDS_SETBUDDYINT) {
				if(containedEdit.GetWindowTextLength() == 0) {
					/* HACK: If the contained edit control is empty and the value wasn't really changed, it's still
					         empty. */
					containedEdit.SetWindowText(TEXT("a"));
					containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
				}
			}
		//} else {
		//	properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS, 0, 0));
		//	containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<SHORT>(properties.currentValue), 0));
		//	if(containedUpDown.GetStyle() & UDS_SETBUDDYINT) {
		//		if(containedEdit.GetWindowTextLength() == 0) {
		//			/* HACK: If the contained edit control is empty and the value wasn't really changed, it's still
		//			         empty. */
		//			containedEdit.SetWindowText(TEXT("a"));
		//			containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<SHORT>(properties.currentValue), 0));
		//		}
		//	}
		//}
	}
	flags.uiActivationPending = FALSE;
	wasHandled = FALSE;
	return 0;
}

LRESULT UpDownTextBox::OnMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_TruncatedText();
	return 0;
}

LRESULT UpDownTextBox::OnSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND hWnd, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(hWnd);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(hWnd, ime);
			}
		}
	}

	if(containedEdit.SendMessage(WM_GETDLGCODE, 0, 0) & DLGC_HASSETSEL) {
		if(GetAsyncKeyState(VK_TAB) & 0x8000) {
			containedEdit.SendMessage(EM_SETSEL, 0, -1);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT UpDownTextBox::OnUpdate(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deBeforeDrawText)) {
		Raise_BeforeDrawText();
	}

	return 0;
}

LRESULT UpDownTextBox::OnSetRedraw(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	containedEdit.SetRedraw(static_cast<BOOL>(wParam));
	containedUpDown.SetRedraw(static_cast<BOOL>(wParam));
	wasHandled = FALSE;
	return 0;
}

LRESULT UpDownTextBox::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT UpDownTextBox::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT UpDownTextBox::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
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

LRESULT UpDownTextBox::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
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

		// size the edit control to 100% of the container
		CRect clientRectangle;
		GetClientRect(&clientRectangle);
		containedEdit.MoveWindow(clientRectangle.left, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height());
		// now reset the buddy, so that the up down control places itself into the edit control
		containedUpDown.SendMessage(UDM_SETBUDDY, reinterpret_cast<WPARAM>(static_cast<HWND>(containedEdit)), 0);
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		containedEdit.Invalidate();
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

LRESULT UpDownTextBox::OnEditChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT UpDownTextBox::OnEditContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y, htTextBox, &showDefaultMenu);
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT UpDownTextBox::OnEditInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = containedEdit.DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if((ime != imeNoControl) && (GetFocus() == containedEdit)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(containedEdit, ime);
	}
	return lr;
}

LRESULT UpDownTextBox::OnEditKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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
	return containedEdit.DefWindowProc(message, wParam, lParam);
}

LRESULT UpDownTextBox::OnEditKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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
	return containedEdit.DefWindowProc(message, wParam, lParam);
}

LRESULT UpDownTextBox::OnEditLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedEdit.MapWindowPoints(*this, &mousePosition, 1);
		Raise_DblClick(button, shift, mousePosition.x, mousePosition.y, htTextBox);
	}

	return 0;
}

LRESULT UpDownTextBox::OnEditLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedEdit.MapWindowPoints(*this, &mousePosition, 1);
		Raise_MDblClick(button, shift, mousePosition.x, mousePosition.y, htTextBox);
	}

	return 0;
}

LRESULT UpDownTextBox::OnEditMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseHover(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus_Edit.lastPosition.x, mouseStatus_Edit.lastPosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedEdit.MapWindowPoints(*this, &mousePosition, 1);
		Raise_MouseMove(button, shift, mousePosition.x, mousePosition.y, htTextBox);
	} else if(!mouseStatus_Edit.enteredControl) {
		mouseStatus_Edit.EnterControl();
	}

	return 0;
}

LRESULT UpDownTextBox::OnEditMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, htTextBox, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus_Edit.enteredControl) {
		mouseStatus_Edit.EnterControl();
	}

	return 0;
}

LRESULT UpDownTextBox::OnEditRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedEdit.MapWindowPoints(*this, &mousePosition, 1);
		Raise_RDblClick(button, shift, mousePosition.x, mousePosition.y, htTextBox);
	}

	return 0;
}

LRESULT UpDownTextBox::OnEditRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControlBase::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControlBase::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT UpDownTextBox::OnEditSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_UPDWNTXTBOX_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = containedEdit.DefWindowProc(message, wParam, lParam);

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
		FireOnChanged(DISPID_UPDWNTXTBOX_FONT);
	}

	return lr;
}

LRESULT UpDownTextBox::OnEditSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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

	return containedEdit.DefWindowProc(message, wParam, lParam);
}

LRESULT UpDownTextBox::OnEditSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_UPDWNTXTBOX_TEXT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = containedEdit.DefWindowProc(message, wParam, lParam);
	SetDirty(TRUE);
	FireOnChanged(DISPID_UPDWNTXTBOX_TEXT);
	SendOnDataChange();
	return lr;
}

LRESULT UpDownTextBox::OnEditXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedEdit.MapWindowPoints(*this, &mousePosition, 1);
		Raise_XDblClick(button, shift, mousePosition.x, mousePosition.y, htTextBox);
	}

	return 0;
}

LRESULT UpDownTextBox::OnEditXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnEditXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedEdit.MapWindowPoints(*this, &mousePosition, 1);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htTextBox);

	return 0;
}

LRESULT UpDownTextBox::OnSetCueBanner(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_UPDWNTXTBOX_CUEBANNER) == S_FALSE) {
		return FALSE;
	}

	LRESULT lr = containedEdit.DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.displayCueBannerOnFocus = wParam;
		properties.cueBanner = reinterpret_cast<LPWSTR>(lParam);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_UPDWNTXTBOX_CUEBANNER);
	SendOnDataChange();
	return lr;
}

LRESULT UpDownTextBox::OnUpDownContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y, htUpDown, &showDefaultMenu);
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT UpDownTextBox::OnUpDownLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return containedUpDown.SendMessage(WM_LBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedUpDown.ClientToScreen(&mousePosition);
		ScreenToClient(&mousePosition);
		Raise_DblClick(button, shift, mousePosition.x, mousePosition.y, htUpDown);
	}

	return 0;
}

LRESULT UpDownTextBox::OnUpDownLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return containedUpDown.SendMessage(WM_MBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedUpDown.ClientToScreen(&mousePosition);
		ScreenToClient(&mousePosition);
		Raise_MDblClick(button, shift, mousePosition.x, mousePosition.y, htUpDown);
	}

	return 0;
}

LRESULT UpDownTextBox::OnUpDownMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseHover(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus_UpDown.lastPosition.x, mouseStatus_UpDown.lastPosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedUpDown.ClientToScreen(&mousePosition);
		ScreenToClient(&mousePosition);
		Raise_MouseMove(button, shift, mousePosition.x, mousePosition.y, htUpDown);
	} else if(!mouseStatus_UpDown.enteredControl) {
		mouseStatus_UpDown.EnterControl();
	}

	return 0;
}

LRESULT UpDownTextBox::OnUpDownMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, htUpDown, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus_UpDown.enteredControl) {
		mouseStatus_UpDown.EnterControl();
	}

	return 0;
}

LRESULT UpDownTextBox::OnUpDownRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return containedUpDown.SendMessage(WM_RBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedUpDown.ClientToScreen(&mousePosition);
		ScreenToClient(&mousePosition);
		Raise_RDblClick(button, shift, mousePosition.x, mousePosition.y, htUpDown);
	}

	return 0;
}

LRESULT UpDownTextBox::OnUpDownRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControlBase::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControlBase::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	containedEdit.SetFocus();
	return lr;
}

LRESULT UpDownTextBox::OnUpDownSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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

	return containedUpDown.DefWindowProc(message, wParam, lParam);
}

LRESULT UpDownTextBox::OnUpDownXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return containedUpDown.SendMessage(WM_XBUTTONDOWN, wParam, lParam);
	}

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
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedUpDown.ClientToScreen(&mousePosition);
		ScreenToClient(&mousePosition);
		Raise_XDblClick(button, shift, mousePosition.x, mousePosition.y, htUpDown);
	}

	return 0;
}

LRESULT UpDownTextBox::OnUpDownXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnUpDownXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	containedUpDown.ClientToScreen(&mousePosition);
	ScreenToClient(&mousePosition);
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y, htUpDown);

	return 0;
}

LRESULT UpDownTextBox::OnSetAccel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IVirtualUpDownAccelerators> pVUpDownAccelsItf = NULL;
	CComObject<VirtualUpDownAccelerators>* pVUpDownAccelsObj = NULL;
	CComObject<VirtualUpDownAccelerators>::CreateInstance(&pVUpDownAccelsObj);
	pVUpDownAccelsObj->AddRef();
	pVUpDownAccelsObj->Attach(static_cast<int>(wParam), reinterpret_cast<LPUDACCEL>(lParam));
	pVUpDownAccelsObj->QueryInterface(IID_IVirtualUpDownAccelerators, reinterpret_cast<LPVOID*>(&pVUpDownAccelsItf));
	pVUpDownAccelsObj->Release();
	Raise_ChangingAccelerators(pVUpDownAccelsItf, &cancel);

	if(cancel == VARIANT_FALSE) {
		lr = containedUpDown.DefWindowProc(message, wParam, lParam);

		CComPtr<IUpDownAccelerators> pUpDownAccelsItf = NULL;
		get_Accelerators(&pUpDownAccelsItf);
		Raise_ChangedAccelerators(pUpDownAccelsItf);
	}

	return lr;
}

LRESULT UpDownTextBox::OnSetRange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = containedUpDown.DefWindowProc(message, wParam, lParam);
	containedUpDown.Invalidate();
	if(properties.automaticallyCorrectValue && (flags.noAutoCorrection == 0)) {
		/*if(IsComctl32Version580OrNewer()) {*/
			properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS32, 0, 0));
			containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
		/*} else {
			properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS, 0, 0));
			containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<WORD>(properties.currentValue), 0));
		}*/
	}
	return lr;
}


void UpDownTextBox::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(containedEdit.IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			containedEdit.SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
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

					containedEdit.SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					containedEdit.SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				containedEdit.SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			containedEdit.Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT UpDownTextBox::Raise_BeforeDrawText(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeforeDrawText();
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_ChangedAccelerators(IUpDownAccelerators* pAccelerators)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedAccelerators(pAccelerators);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_ChangingAccelerators(IVirtualUpDownAccelerators* pAccelerators, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangingAccelerators(pAccelerators, pCancel);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
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

		return Fire_ContextMenu(button, shift, x, y, hitTestDetails, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(containedEdit) == S_OK);
		ATLVERIFY(RevokeDragDrop(containedUpDown) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	switch(hitTestDetails) {
		case htTextBox:
			pMouseStatusToUse = &mouseStatus_Edit;
			break;
		case htUpDown:
			pMouseStatusToUse = &mouseStatus_UpDown;
			break;
		default:
			ATLASSERT((hitTestDetails == htTextBox) || (hitTestDetails == htUpDown));
			return E_INVALIDARG;
			break;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!pMouseStatusToUse->enteredControl) {
			Raise_MouseEnter(button, shift, x, y, hitTestDetails);
		}
		if(!pMouseStatusToUse->hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			trackingOptions.hwndTrack = containedEdit;
			TrackMouseEvent(&trackingOptions);
			trackingOptions.hwndTrack = containedUpDown;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y, hitTestDetails);
		}
		pMouseStatusToUse->StoreClickCandidate(button);
		switch(hitTestDetails) {
			case htTextBox:
				containedEdit.SetCapture();
				break;
			case htUpDown:
				containedUpDown.SetCapture();
				break;
		}

		return Fire_MouseDown(button, shift, x, y, hitTestDetails);
	} else {
		if(!pMouseStatusToUse->enteredControl) {
			pMouseStatusToUse->EnterControl();
		}
		if(!pMouseStatusToUse->hoveredControl) {
			pMouseStatusToUse->HoverControl();
		}
		pMouseStatusToUse->StoreClickCandidate(button);
		switch(hitTestDetails) {
			case htTextBox:
				containedEdit.SetCapture();
				break;
			case htUpDown:
				containedUpDown.SetCapture();
				break;
		}
		return S_OK;
	}
}

inline HRESULT UpDownTextBox::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	MouseStatus* pMouseStatusToUse = NULL;
	switch(hitTestDetails) {
		case htTextBox:
			trackingOptions.hwndTrack = containedEdit;
			pMouseStatusToUse = &mouseStatus_Edit;
			break;
		case htUpDown:
			trackingOptions.hwndTrack = containedUpDown;
			pMouseStatusToUse = &mouseStatus_UpDown;
			break;
		default:
			ATLASSERT((hitTestDetails == htTextBox) || (hitTestDetails == htUpDown));
			return E_INVALIDARG;
			break;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	pMouseStatusToUse->EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	switch(hitTestDetails) {
		case htTextBox:
			mouseStatus_Edit.HoverControl();
			break;
		case htUpDown:
			mouseStatus_UpDown.HoverControl();
			break;
		default:
			ATLASSERT((hitTestDetails == htTextBox) || (hitTestDetails == htUpDown));
			break;
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT UpDownTextBox::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	switch(hitTestDetails) {
		case htTextBox:
			mouseStatus_Edit.LeaveControl();
			break;
		case htUpDown:
			mouseStatus_UpDown.LeaveControl();
			break;
		default:
			ATLASSERT((hitTestDetails == htTextBox) || (hitTestDetails == htUpDown));
			break;
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT UpDownTextBox::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	MouseStatus* pMouseStatusToUse = NULL;
	switch(hitTestDetails) {
		case htTextBox:
			pMouseStatusToUse = &mouseStatus_Edit;
			break;
		case htUpDown:
			pMouseStatusToUse = &mouseStatus_UpDown;
			break;
		default:
			ATLASSERT((hitTestDetails == htTextBox) || (hitTestDetails == htUpDown));
			return E_INVALIDARG;
			break;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(!pMouseStatusToUse->enteredControl) {
		Raise_MouseEnter(button, shift, x, y, hitTestDetails);
	}
	pMouseStatusToUse->lastPosition.x = x;
	pMouseStatusToUse->lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y, hitTestDetails);
	}

	MouseStatus* pMouseStatusToUse = NULL;
	switch(hitTestDetails) {
		case htTextBox:
			pMouseStatusToUse = &mouseStatus_Edit;
			break;
		case htUpDown:
			pMouseStatusToUse = &mouseStatus_UpDown;
			break;
		default:
			ATLASSERT((hitTestDetails == htTextBox) || (hitTestDetails == htUpDown));
			return E_INVALIDARG;
			break;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(pMouseStatusToUse->IsClickCandidate(button)) {
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
			HWND hWindowBelowMouse = WindowFromPoint(cursorPosition);
			if((hWindowBelowMouse != *this) && !IsChild(hWindowBelowMouse)) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		switch(hitTestDetails) {
			case htTextBox:
				if(GetCapture() == containedEdit) {
					ReleaseCapture();
				}
				break;
			case htUpDown:
				if(GetCapture() == containedUpDown) {
					ReleaseCapture();
				}
				break;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y, hitTestDetails);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y, hitTestDetails);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y, hitTestDetails);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y, hitTestDetails);
					}
					break;
			}
		}

		pMouseStatusToUse->RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT UpDownTextBox::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	MouseStatus* pMouseStatusToUse = NULL;
	switch(hitTestDetails) {
		case htTextBox:
			pMouseStatusToUse = &mouseStatus_Edit;
			break;
		case htUpDown:
			pMouseStatusToUse = &mouseStatus_UpDown;
			break;
		default:
			ATLASSERT((hitTestDetails == htTextBox) || (hitTestDetails == htUpDown));
			return E_INVALIDARG;
			break;
	}
	ATLASSERT_POINTER(pMouseStatusToUse, MouseStatus);

	if(!pMouseStatusToUse->enteredControl) {
		Raise_MouseEnter(button, shift, x, y, hitTestDetails);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		POINT buffer = {mousePosition.x, mousePosition.y};
		HitTestConstants hitTestDetails = ((WindowFromPoint(buffer) == containedEdit) ? htTextBox : htUpDown);

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
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT UpDownTextBox::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	HitTestConstants hitTestDetails = ((WindowFromPoint(buffer) == containedEdit) ? htTextBox : htUpDown);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
		}
	//}
	return S_OK;
}

inline HRESULT UpDownTextBox::Raise_OLEDragLeave(void)
{
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		POINT buffer = {dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y};
		HitTestConstants hitTestDetails = ((WindowFromPoint(buffer) == containedEdit) ? htTextBox : htUpDown);

		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, hitTestDetails);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT UpDownTextBox::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	POINT buffer = {mousePosition.x, mousePosition.y};
	HitTestConstants hitTestDetails = ((WindowFromPoint(buffer) == containedEdit) ? htTextBox : htUpDown);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
		}
	//}
	return S_OK;
}

inline HRESULT UpDownTextBox::Raise_OutOfMemory(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OutOfMemory();
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the contained controls
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(containedEdit, static_cast<IDropTarget*>(this)) == S_OK);
		ATLVERIFY(RegisterDragDrop(containedUpDown, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_TextChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TextChanged();
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_TruncatedText(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TruncatedText();
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_ValueChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ValueChanged();
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_ValueChanging(LONG currentValue, LONG delta, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ValueChanging(currentValue, delta, pCancel);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_WritingDirectionChanged(WritingDirectionConstants newWritingDirection)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_WritingDirectionChanged(newWritingDirection);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT UpDownTextBox::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}


void UpDownTextBox::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		if(containedEdit.IsWindow()) {
			containedEdit.GetWindowText(&properties.text);
		}
		if(containedUpDown.IsWindow()) {
			/*if(IsComctl32Version580OrNewer()) {*/
				properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS32, 0, 0));
			/*} else {
				properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS, 0, 0));
			}*/
		}

		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

IMEModeConstants UpDownTextBox::GetCurrentIMEContextMode(HWND hWnd)
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

void UpDownTextBox::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
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

IMEModeConstants UpDownTextBox::GetEffectiveIMEMode(void)
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

DWORD UpDownTextBox::GetEditExStyleBits(void)
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

DWORD UpDownTextBox::GetEditStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
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
	return style;
}

DWORD UpDownTextBox::GetUpDownExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD UpDownTextBox::GetUpDownStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_VISIBLE | UDS_AUTOBUDDY;
	if(properties.automaticallySetText) {
		style |= UDS_SETBUDDYINT;
	}
	if(!properties.groupDigits) {
		style |= UDS_NOTHOUSANDS;
	}
	if(properties.hotTracking) {
		style |= UDS_HOTTRACK;
	}
	switch(properties.orientation) {
		case oHorizontal:
			style |= UDS_HORZ;
			break;
	}
	if(properties.processArrowKeys) {
		style |= UDS_ARROWKEYS;
	}
	switch(properties.upDownPosition) {
		case udLeftOfTextBox:
			style |= UDS_ALIGNLEFT;
			break;
		case udRightOfTextBox:
			style |= UDS_ALIGNRIGHT;
			break;
	}
	if(properties.wrapAtBoundaries) {
		style |= UDS_WRAP;
	}
	return style;
}

void UpDownTextBox::RecreateEditWindow(void)
{
	if(containedEdit.IsWindow()) {
		containedEdit.GetWindowText(&properties.text);

		containedEdit.DestroyWindow();

		CRect clientRectangle;
		GetClientRect(&clientRectangle);
		HWND hWndEdit = CreateWindowEx(GetEditExStyleBits(), WC_EDIT, NULL, GetEditStyleBits(), clientRectangle.left, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height(), *this, NULL, ModuleHelper::GetModuleInstance(), NULL);
		containedEdit.SubclassWindow(hWndEdit);

		/* TODO: Why is this necessary? If we're using RTL features, we sometimes end up with an edit control
		         that has this style set. Why? */
		containedEdit.ModifyStyleEx(WS_EX_RIGHT, 0);
		containedEdit.ModifyStyle(ES_RIGHT, 0);

		// now reset the buddy, so that the up down control places itself into the edit control
		containedUpDown.SendMessage(UDM_SETBUDDY, reinterpret_cast<WPARAM>(static_cast<HWND>(containedEdit)), 0);

		SendEditConfigurationMessages();
	}
}

void UpDownTextBox::RecreateUpDownWindow(void)
{
	if(containedUpDown.IsWindow()) {
		/*if(IsComctl32Version580OrNewer()) {*/
			properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS32, 0, 0));
		/*} else {
			properties.currentValue = static_cast<LONG>(containedUpDown.SendMessage(UDM_GETPOS, 0, 0));
		}*/

		containedUpDown.DestroyWindow();

		CRect clientRectangle;
		GetClientRect(&clientRectangle);
		HWND hWndUD = CreateWindowEx(GetUpDownExStyleBits(), UPDOWN_CLASS, NULL, GetUpDownStyleBits(), clientRectangle.right, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height(), *this, NULL, ModuleHelper::GetModuleInstance(), NULL);
		containedUpDown.SubclassWindow(hWndUD);

		// size the edit control to 100% of the container
		containedEdit.MoveWindow(clientRectangle.left, clientRectangle.top, clientRectangle.Width(), clientRectangle.Height());
		// now reset the buddy, so that the up down control places itself into the edit control
		containedUpDown.SendMessage(UDM_SETBUDDY, reinterpret_cast<WPARAM>(static_cast<HWND>(containedEdit)), 0);

		SendUpDownConfigurationMessages();
	}
}

void UpDownTextBox::SendConfigurationMessages(void)
{
	// the up down control seems to override the edit control's alignment settings
	switch(properties.hAlignment) {
		case halLeft:
			containedEdit.ModifyStyle(ES_CENTER | ES_RIGHT, ES_LEFT);
			break;
		case halCenter:
			containedEdit.ModifyStyle(ES_LEFT | ES_RIGHT, ES_CENTER);
			break;
		case halRight:
			containedEdit.ModifyStyle(ES_LEFT | ES_CENTER, ES_RIGHT);
			break;
	}

	DWORD flags = static_cast<DWORD>(containedEdit.SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0));
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
	containedEdit.SendMessage(EM_SETIMESTATUS, EMSIS_COMPOSITIONSTRING, flags);
	containedEdit.SendMessage(EM_SETLIMITTEXT, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength), 0);
	containedEdit.SendMessage(EM_SETREADONLY, properties.readOnlyTextBox, 0);

	switch(properties.base) {
		case bDecimal:
			containedUpDown.SendMessage(UDM_SETBASE, 10, 0);
			break;
		case bHexadecimal:
			containedUpDown.SendMessage(UDM_SETBASE, 16, 0);
			break;
	}
	++this->flags.noAutoCorrection;
	containedUpDown.SendMessage(UDM_SETRANGE32, properties.minimum, properties.maximum);
	--this->flags.noAutoCorrection;
	/*if(IsComctl32Version580OrNewer()) {*/
		containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
	/*} else {
		containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<WORD>(properties.currentValue), 0));
	}*/

	if(!(containedUpDown.GetStyle() & UDS_SETBUDDYINT)) {
		containedEdit.SetWindowText(COLE2CT(properties.text));
	}
	if(RunTimeHelper::IsCommCtrl6()) {
		containedEdit.SendMessage(EM_SETCUEBANNER, properties.displayCueBannerOnFocus, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
	}
	containedEdit.SendMessage(EM_SETMODIFY, properties.modified, 0);

	ApplyFont();
	containedEdit.SendMessage(EM_SETMARGINS, (EC_LEFTMARGIN | EC_RIGHTMARGIN), MAKELPARAM((properties.leftMargin == -1 ? EC_USEFONTINFO : properties.leftMargin), (properties.rightMargin == -1 ? EC_USEFONTINFO : properties.rightMargin)));
}

void UpDownTextBox::SendEditConfigurationMessages(void)
{
	// the up down control seems to override the edit control's alignment settings
	switch(properties.hAlignment) {
		case halLeft:
			containedEdit.ModifyStyle(ES_CENTER | ES_RIGHT, ES_LEFT);
			break;
		case halCenter:
			containedEdit.ModifyStyle(ES_LEFT | ES_RIGHT, ES_CENTER);
			break;
		case halRight:
			containedEdit.ModifyStyle(ES_LEFT | ES_CENTER, ES_RIGHT);
			break;
	}

	DWORD flags = static_cast<DWORD>(containedEdit.SendMessage(EM_GETIMESTATUS, EMSIS_COMPOSITIONSTRING, 0));
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
	containedEdit.SendMessage(EM_SETIMESTATUS, EMSIS_COMPOSITIONSTRING, flags);
	containedEdit.SendMessage(EM_SETLIMITTEXT, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength), 0);
	containedEdit.SendMessage(EM_SETREADONLY, properties.readOnlyTextBox, 0);

	containedEdit.SetWindowText(COLE2CT(properties.text));
	if(RunTimeHelper::IsCommCtrl6()) {
		containedEdit.SendMessage(EM_SETCUEBANNER, 0, reinterpret_cast<LPARAM>(OLE2W(properties.cueBanner)));
	}
	containedEdit.SendMessage(EM_SETMODIFY, properties.modified, 0);

	ApplyFont();
	containedEdit.SendMessage(EM_SETMARGINS, (EC_LEFTMARGIN | EC_RIGHTMARGIN), MAKELPARAM((properties.leftMargin == -1 ? EC_USEFONTINFO : properties.leftMargin), (properties.rightMargin == -1 ? EC_USEFONTINFO : properties.rightMargin)));
}

void UpDownTextBox::SendUpDownConfigurationMessages(void)
{
	switch(properties.base) {
		case bDecimal:
			containedUpDown.SendMessage(UDM_SETBASE, 10, 0);
			break;
		case bHexadecimal:
			containedUpDown.SendMessage(UDM_SETBASE, 16, 0);
			break;
	}
	++flags.noAutoCorrection;
	containedUpDown.SendMessage(UDM_SETRANGE32, properties.minimum, properties.maximum);
	--flags.noAutoCorrection;
	/*if(IsComctl32Version580OrNewer()) {*/
		containedUpDown.SendMessage(UDM_SETPOS32, 0, properties.currentValue);
	/*} else {
		containedUpDown.SendMessage(UDM_SETPOS, 0, MAKELPARAM(static_cast<WORD>(properties.currentValue), 0));
	}*/

	if(!(containedUpDown.GetStyle() & UDS_SETBUDDYINT)) {
		containedEdit.SetWindowText(COLE2CT(properties.text));
	}
}

HCURSOR UpDownTextBox::MousePointerConst2hCursor(MousePointerConstants mousePointer)
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

BOOL UpDownTextBox::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}