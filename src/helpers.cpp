// helpers.cpp: global helper functions, macros and other stuff

#include "stdafx.h"
#include "helpers.h"
#include "HotKeyBox.h"
#include "IPAddressBox.h"
#include "TextBox.h"
#include "UpDownTextBox.h"


void OLEKEYSTATE2BUTTONANDSHIFT(int bitField, SHORT& mouseButtons, SHORT& shiftButtons)
{
	mouseButtons = shiftButtons = 0;
	if(bitField & MK_SHIFT) {
		shiftButtons |= 1 /*ShiftConstants.vbShiftMask*/;
	}
	if(bitField & MK_CONTROL) {
		shiftButtons |= 2 /*ShiftConstants.vbCtrlMask*/;
	}
	if(bitField & MK_ALT) {
		shiftButtons |= 4 /*ShiftConstants.vbAltMask*/;
	}
	if(bitField & MK_LBUTTON) {
		mouseButtons |= 1 /*MouseButtonConstants.vbLeftButton*/;
	}
	if(bitField & MK_MBUTTON) {
		mouseButtons |= 4 /*MouseButtonConstants.vbMiddleButton*/;
	}
	if(bitField & MK_RBUTTON) {
		mouseButtons |= 2 /*MouseButtonConstants.vbRightButton*/;
	}
}

void WPARAM2BUTTONANDSHIFT(int bitField, SHORT& mouseButtons, SHORT& shiftButtons)
{
	mouseButtons = shiftButtons = 0;
	if(bitField == -1) {
		int swappedButtons = GetSystemMetrics(SM_SWAPBUTTON);
		if(GetAsyncKeyState(VK_SHIFT) & 0x8000) {
			shiftButtons |= 1 /*ShiftConstants.vbShiftMask*/;
		}
		if(GetAsyncKeyState(VK_CONTROL) & 0x8000) {
			shiftButtons |= 2 /*ShiftConstants.vbCtrlMask*/;
		}
		if(GetAsyncKeyState(VK_MENU) & 0x8000) {
			shiftButtons |= 4 /*ShiftConstants.vbAltMask*/;
		}
		if(GetAsyncKeyState(swappedButtons ? VK_RBUTTON : VK_LBUTTON) & 0x8000) {
			mouseButtons |= 1 /*MouseButtonConstants.vbLeftButton*/;
		}
		if(GetAsyncKeyState(VK_MBUTTON) & 0x8000) {
			mouseButtons |= 4 /*MouseButtonConstants.vbMiddleButton*/;
		}
		if(GetAsyncKeyState(swappedButtons ? VK_LBUTTON : VK_RBUTTON) & 0x8000) {
			mouseButtons |= 2 /*MouseButtonConstants.vbRightButton*/;
		}
		if(GetAsyncKeyState(VK_XBUTTON1) & 0x8000) {
			mouseButtons |= embXButton1;
		}
		if(GetAsyncKeyState(VK_XBUTTON2) & 0x8000) {
			mouseButtons |= embXButton2;
		}
	} else {
		if(bitField & MK_SHIFT) {
			shiftButtons |= 1 /*ShiftConstants.vbShiftMask*/;
		}
		if(bitField & MK_CONTROL) {
			shiftButtons |= 2 /*ShiftConstants.vbCtrlMask*/;
		}
		// NOTE: GetKeyState is right! See documentation of WM_LBUTTONDOWN.
		if(GetKeyState(VK_MENU) & 0x8000) {
			shiftButtons |= 4 /*ShiftConstants.vbAltMask*/;
		}
		if(bitField & MK_LBUTTON) {
			mouseButtons |= 1 /*MouseButtonConstants.vbLeftButton*/;
		}
		if(bitField & MK_MBUTTON) {
			mouseButtons |= 4 /*MouseButtonConstants.vbMiddleButton*/;
		}
		if(bitField & MK_RBUTTON) {
			mouseButtons |= 2 /*MouseButtonConstants.vbRightButton*/;
		}
		if(bitField & MK_XBUTTON1) {
			mouseButtons |= embXButton1;
		}
		if(bitField & MK_XBUTTON2) {
			mouseButtons |= embXButton2;
		}
	}
}

LPTSTR ConvertIntegerToString(char value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_itot_s(static_cast<int>(value), pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(int value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_itot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(long value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ltot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(__int64 value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_i64tot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned char value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ultot_s(static_cast<unsigned long>(value), pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned int value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ultot_s(static_cast<unsigned long>(value), pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned long value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ultot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned __int64 value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ui64tot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

void ConvertStringToInteger(LPTSTR str, char& value)
{
	value = static_cast<char>(_ttoi(str) & 0xFF);
}

void ConvertStringToInteger(LPTSTR str, int& value)
{
	value = _ttoi(str);
}

void ConvertStringToInteger(LPTSTR str, long& value)
{
	value = _ttol(str);
}

void ConvertStringToInteger(LPTSTR str, __int64& value)
{
	value = _ttoi64(str);
}

long ConvertPixelsToTwips(HDC hDC, long pixels, BOOL vertical)
{
	const int twipsPerInch = 1440;
	int pixelsPerInch = 0;
	if(vertical) {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
	} else {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX);
	}
	return (pixels * twipsPerInch) / pixelsPerInch;
}

POINT ConvertPixelsToTwips(HDC hDC, POINT &pixels)
{
	POINT converted = {0};
	converted.x = ConvertPixelsToTwips(hDC, pixels.x, FALSE);
	converted.y = ConvertPixelsToTwips(hDC, pixels.y, TRUE);
	return converted;
}

RECT ConvertPixelsToTwips(HDC hDC, RECT &pixels)
{
	RECT converted = {0};
	converted.bottom = ConvertPixelsToTwips(hDC, pixels.bottom, TRUE);
	converted.left = ConvertPixelsToTwips(hDC, pixels.left, FALSE);
	converted.right = ConvertPixelsToTwips(hDC, pixels.right, FALSE);
	converted.top = ConvertPixelsToTwips(hDC, pixels.top, TRUE);
	return converted;
}

SIZE ConvertPixelsToTwips(HDC hDC, SIZE &pixels)
{
	SIZE converted = {0};
	converted.cx = ConvertPixelsToTwips(hDC, pixels.cx, FALSE);
	converted.cy = ConvertPixelsToTwips(hDC, pixels.cy, TRUE);
	return converted;
}

long ConvertTwipsToPixels(HDC hDC, long twips, BOOL vertical)
{
	const int twipsPerInch = 1440;
	int pixelsPerInch = 0;
	if(vertical) {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
	} else {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX);
	}
	return (twips * pixelsPerInch) / twipsPerInch;
}

POINT ConvertTwipsToPixels(HDC hDC, POINT &twips)
{
	POINT converted = {0};
	converted.x = ConvertTwipsToPixels(hDC, twips.x, FALSE);
	converted.y = ConvertTwipsToPixels(hDC, twips.y, TRUE);
	return converted;
}

RECT ConvertTwipsToPixels(HDC hDC, RECT &twips)
{
	RECT converted = {0};
	converted.bottom = ConvertTwipsToPixels(hDC, twips.bottom, TRUE);
	converted.left = ConvertTwipsToPixels(hDC, twips.left, FALSE);
	converted.right = ConvertTwipsToPixels(hDC, twips.right, FALSE);
	converted.top = ConvertTwipsToPixels(hDC, twips.top, TRUE);
	return converted;
}

SIZE ConvertTwipsToPixels(HDC hDC, SIZE &twips)
{
	SIZE converted = {0};
	converted.cx = ConvertTwipsToPixels(hDC, twips.cx, FALSE);
	converted.cy = ConvertTwipsToPixels(hDC, twips.cy, TRUE);
	return converted;
}

COLORREF OLECOLOR2COLORREF(OLE_COLOR oleColor, HPALETTE hPalette/* = NULL*/)
{
	COLORREF color = RGB(0x00, 0x00, 0x00);
	OleTranslateColor(oleColor, hPalette, &color);
	return color;
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, LPTSTR description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	// convert the description string to an OLE string
	LPOLESTR oleDescription = NULL;
	if(description) {
		oleDescription = SysAllocString(CT2COLE(description));
	} else {
		if(HRESULT_FACILITY(hError) == FACILITY_WIN32) {
			// get the error from the system
			LPTSTR win32Description = NULL;
			if(!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER, NULL, HRESULT_CODE(hError), MAKELANGID(LANG_USER_DEFAULT, SUBLANG_DEFAULT), reinterpret_cast<LPTSTR>(&win32Description), 0, NULL)) {
				return HRESULT_FROM_WIN32(GetLastError());
			}

			// convert the multibyte string to an OLE string
			if(win32Description) {
				oleDescription = SysAllocString(CT2COLE(win32Description));
				LocalFree(win32Description);
			}
		}
	}

	// convert the title string to an OLE string
	LPOLESTR oleTitle = NULL;
	if(title) {
		oleTitle = SysAllocString(CT2COLE(title));
	}

	// convert the help filename string to an OLE string
	LPOLESTR oleHelpFileName = NULL;
	if(helpFileName) {
		oleHelpFileName = SysAllocString(CT2COLE(helpFileName));
	}

	// retrieve the ICreateErrorInfo interface
	LPCREATEERRORINFO pErrorInfoCreator = NULL;
	ATLVERIFY(SUCCEEDED(CreateErrorInfo(&pErrorInfoCreator)));

	// fill the error information into it
	pErrorInfoCreator->SetGUID(source);
	if(oleDescription) {
		pErrorInfoCreator->SetDescription(oleDescription);
	}
	if(oleTitle) {
		pErrorInfoCreator->SetSource(oleTitle);
	}
	if(oleHelpFileName) {
		pErrorInfoCreator->SetHelpFile(oleHelpFileName);
	}
	pErrorInfoCreator->SetHelpContext(helpContext);

	SysFreeString(oleDescription);
	SysFreeString(oleTitle);
	SysFreeString(oleHelpFileName);

	// retrieve the IErrorInfo interface
	LPERRORINFO pErrorInfo = NULL;
	HRESULT hSuccess = pErrorInfoCreator->QueryInterface(IID_IErrorInfo, reinterpret_cast<LPVOID*>(&pErrorInfo));
	if(FAILED(hSuccess)) {
		pErrorInfoCreator->Release();
		return hSuccess;
	}

	// set this error information in the current thread
	hSuccess = SetErrorInfo(0, pErrorInfo);
	if(FAILED(hSuccess)) {
		pErrorInfoCreator->Release();
		pErrorInfo->Release();
		return hSuccess;
	}

	// finally release the interfaces
	pErrorInfoCreator->Release();
	pErrorInfo->Release();

	// return the error code that was asked to be dispatched
	return hError;
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, UINT description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(hError, source, title, COLE2T(GetResString(description)), helpContext, helpFileName);
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, LPTSTR description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(hError, source, COLE2T(GetResString(title)), description, helpContext, helpFileName);
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, UINT description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(hError, source, COLE2T(GetResString(title)), COLE2T(GetResString(description)), helpContext, helpFileName);
}

HRESULT DispatchError(DWORD errorCode, REFCLSID source, LPTSTR title, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	// dispatch the requested error message
	return DispatchError(HRESULT_FROM_WIN32(errorCode), source, title, static_cast<LPTSTR>(NULL), helpContext, helpFileName);
}

HRESULT DispatchError(DWORD errorCode, REFCLSID source, UINT title, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(errorCode, source, COLE2T(GetResString(title)), helpContext, helpFileName);
}


HRESULT ReadVariantFromStream(LPSTREAM pStream, VARTYPE expectedVarType, LPVARIANT pVariant)
{
	ATLASSERT(expectedVarType == VT_BOOL || expectedVarType == VT_DATE || expectedVarType == VT_I2 || expectedVarType == VT_I4);

	HRESULT hr = pStream->Read(&pVariant->vt, sizeof(VARTYPE), NULL);
	if(SUCCEEDED(hr)) {
		if(pVariant->vt != expectedVarType) {
			return DISP_E_TYPEMISMATCH;
		}
		switch(pVariant->vt) {
			case VT_BOOL:
				hr = pStream->Read(&pVariant->boolVal, sizeof(VARIANT_BOOL), NULL);
				break;
			case VT_DATE:
				hr = pStream->Read(&pVariant->date, sizeof(DATE), NULL);
				break;
			case VT_I2:
				hr = pStream->Read(&pVariant->iVal, sizeof(SHORT), NULL);
				break;
			case VT_I4:
				hr = pStream->Read(&pVariant->lVal, sizeof(LONG), NULL);
				break;
		}
	}
	return hr;
}

HRESULT ReadVariantFromStream_Legacy(LPSTREAM pStream, VARTYPE expectedVarType, LPVARIANT pVariant)
{
	ATLASSERT(expectedVarType == VT_BOOL || expectedVarType == VT_DATE || expectedVarType == VT_I2 || expectedVarType == VT_I4);

	HRESULT hr = pStream->Read(pVariant, sizeof(VARIANT), NULL);
	if(SUCCEEDED(hr)) {
		if(pVariant->vt != expectedVarType) {
			return DISP_E_TYPEMISMATCH;
		}
	}
	return hr;
}

HRESULT WriteVariantToStream(LPSTREAM pStream, LPVARIANT pVariant)
{
	ATLASSERT(pVariant->vt == VT_BOOL || pVariant->vt == VT_DATE || pVariant->vt == VT_I2 || pVariant->vt == VT_I4);

	HRESULT hr = pStream->Write(&pVariant->vt, sizeof(VARTYPE), NULL);
	if(SUCCEEDED(hr)) {
		switch(pVariant->vt) {
			case VT_BOOL:
				hr = pStream->Write(&pVariant->boolVal, sizeof(VARIANT_BOOL), NULL);
				break;
			case VT_DATE:
				hr = pStream->Write(&pVariant->date, sizeof(DATE), NULL);
				break;
			case VT_I2:
				hr = pStream->Write(&pVariant->iVal, sizeof(SHORT), NULL);
				break;
			case VT_I4:
				hr = pStream->Write(&pVariant->lVal, sizeof(LONG), NULL);
				break;
		}
	}
	return hr;
}

WPARAM GetButtonStateBitField(void)
{
	WPARAM bitField = NULL;

	if(GetAsyncKeyState(VK_CONTROL) & 0x8000) {
		bitField |= MK_CONTROL;
	}
	if(GetAsyncKeyState(VK_SHIFT) & 0x8000) {
		bitField |= MK_SHIFT;
	}
	if(GetAsyncKeyState(VK_LBUTTON) & 0x8000) {
		bitField |= MK_LBUTTON;
	}
	if(GetAsyncKeyState(VK_MBUTTON) & 0x8000) {
		bitField |= MK_MBUTTON;
	}
	if(GetAsyncKeyState(VK_XBUTTON1) & 0x8000) {
		bitField |= MK_XBUTTON1;
	}
	if(GetAsyncKeyState(VK_XBUTTON2) & 0x8000) {
		bitField |= MK_XBUTTON2;
	}

	// according to MSDN these two flags are never set
	if(GetAsyncKeyState(VK_MENU) & 0x8000) {
		bitField |= MK_ALT;
	}
	if(GetAsyncKeyState(VK_RBUTTON) & 0x8000) {
		bitField |= MK_RBUTTON;
	}

	return bitField;
}

DllVersion GetDllVersion(LPCTSTR dllName)
{
	DllVersion version;
	version.majorVersion = static_cast<DWORD>(-1);

	// load the DLL
	//HMODULE hMod = LoadLibrary(dllName);
	HMODULE hMod = GetModuleHandle(dllName);
	if(hMod) {
		DLLGETVERSIONPROC pDllGetVersion = reinterpret_cast<DLLGETVERSIONPROC>(GetProcAddress(hMod, "DllGetVersion"));
		if(pDllGetVersion) {
			DLLVERSIONINFO versionInfo = {0};
			versionInfo.cbSize = sizeof(versionInfo);

			HRESULT hr = pDllGetVersion(&versionInfo);
			if(SUCCEEDED(hr)) {
				version.majorVersion = versionInfo.dwMajorVersion;
				version.minorVersion = versionInfo.dwMinorVersion;
				version.buildNumber = versionInfo.dwBuildNumber;
				switch(versionInfo.dwPlatformID) {
					case DLLVER_PLATFORM_WINDOWS:
						version.targetPlatform = DllVersion::Windows;
						break;
					case DLLVER_PLATFORM_NT:
						version.targetPlatform = DllVersion::WindowsNT;
						break;
					default:
						version.targetPlatform = DllVersion::Unknown;
						break;
				}
			}
		}
		//FreeLibrary(hMod);
	}

	return version;
}

CComBSTR GetResString(UINT stringToLoad)
{
	CComBSTR comString;
	comString.LoadString(stringToLoad);
	return comString;
}

CComBSTR GetResStringWithNumber(UINT stringToLoad, int numberToInsert)
{
	CComBSTR result = GetResString(stringToLoad);
	int arraySize = result.Length() + 50;
	LPTSTR pBuffer = new TCHAR[arraySize];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, arraySize * sizeof(TCHAR), COLE2CT(result), numberToInsert)));
	result = CComBSTR(pBuffer);
	delete[] pBuffer;
	return result;
}

HBITMAP LoadJPGResource(UINT jpgToLoad)
{
	HRSRC hResource = FindResource(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(jpgToLoad), TEXT("JPG"));
	if(hResource) {
		HGLOBAL hResourceMem = LoadResource(ModuleHelper::GetResourceInstance(), hResource);
		if(hResourceMem) {
			LPVOID pResourceData = LockResource(hResourceMem);
			if(pResourceData) {
				DWORD resourceSize = SizeofResource(ModuleHelper::GetResourceInstance(), hResource);
				HGLOBAL hJPGMem = GlobalAlloc(GMEM_MOVEABLE, resourceSize);
				if(hJPGMem) {
					LPVOID pJPGData = GlobalLock(hJPGMem);
					if(pJPGData) {
						CopyMemory(pJPGData, pResourceData, resourceSize);

						CComPtr<IStream> pStream;
						if(SUCCEEDED(CreateStreamOnHGlobal(hJPGMem, TRUE, &pStream))) {
							CComPtr<IPictureDisp> pPictureDisp;
							if(SUCCEEDED(OleLoadPicture(pStream, 0, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&pPictureDisp)))) {
								CComQIPtr<IPicture, &IID_IPicture> pPicture(pPictureDisp);
								if(pPicture) {
									HBITMAP hImage = NULL;
									pPicture->get_Handle(reinterpret_cast<OLE_HANDLE*>(&hImage));
									if(hImage) {
										return CopyBitmap(hImage);
									}
								}
							}
						}
						GlobalUnlock(hJPGMem);
					}
					GlobalFree(hJPGMem);
				}
			}
		}
	}
	return NULL;
}

BOOL IsValidAcceleratorIndex(int acceleratorIndex, UpDownTextBox* pUpDownTextBox)
{
	if(acceleratorIndex < 0) {
		return FALSE;
	}

	HWND hWndUpDown = NULL;
	if(pUpDownTextBox) {
		OLE_HANDLE h = NULL;
		pUpDownTextBox->get_hWndUpDown(&h);
		hWndUpDown = static_cast<HWND>(LongToHandle(h));
	}

	return IsValidAcceleratorIndex(acceleratorIndex, hWndUpDown);
}

BOOL IsValidAcceleratorIndex(int acceleratorIndex, HWND hWndUpDown)
{
	if(acceleratorIndex < 0 || !IsWindow(hWndUpDown)) {
		return FALSE;
	}

	int accelerators = static_cast<int>(SendMessage(hWndUpDown, UDM_GETACCEL, 0, 0));
	return (acceleratorIndex < accelerators);
}

int LoadLocalizedStringResource(UINT stringToLoad, LPTSTR buffer, int bufferSize)
{
	ZeroMemory(buffer, bufferSize);

	// retrieve the correct locale
	LCID threadLocale = GetThreadLocale();
	LANGID commonControlsLanguage = GetMUILanguage();
	LCID localeToUse = threadLocale;
	if(commonControlsLanguage == MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL)) {
		// MUI isn't initialized - use system default
		LANGID threadLanguage = LANGIDFROMLCID(threadLocale);
		LANGID languageToUse = threadLanguage;
		/* TODO: On Windows 95/98/NT4 the following code will fail. In this case we should use the language
		         of the dialog resources in shell32. */
		HMODULE hMod = GetModuleHandle(TEXT("kernel32"));
		if(hMod) {
			typedef LANGID WINAPI GetUserDefaultUILanguageFn(void);
			GetUserDefaultUILanguageFn* pfnGetUserDefaultUILanguage = reinterpret_cast<GetUserDefaultUILanguageFn*>(GetProcAddress(hMod, "GetUserDefaultUILanguage"));
			if(pfnGetUserDefaultUILanguage) {
				languageToUse = pfnGetUserDefaultUILanguage();
			}
		}

		localeToUse = (languageToUse == threadLanguage ? threadLocale : MAKELCID(languageToUse, SORT_DEFAULT));
	} else {
		localeToUse = MAKELCID(commonControlsLanguage, SORT_DEFAULT);
	}
	LANGID languageToUse = LANGIDFROMLCID(localeToUse);

	// now load the string
	HMODULE hMod = GetModuleHandle(TEXT("comctl32"));
	// string tables are broken up into "bundles" of 16 strings each
	HRSRC hResource = FindResourceEx(hMod, RT_STRING, reinterpret_cast<LPCTSTR>(static_cast<LONG_PTR>(1 + static_cast<USHORT>(stringToLoad / 16))), languageToUse);
	if(!hResource && (languageToUse != MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL))) {
		languageToUse = MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL);
		hResource = FindResourceEx(hMod, RT_STRING, reinterpret_cast<LPCTSTR>(static_cast<LONG_PTR>(1 + static_cast<USHORT>(stringToLoad / 16))), languageToUse);
	}

	int stringLengthToCopy = 0;
	if(hResource) {
		PWCHAR pData = reinterpret_cast<PWCHAR>(LoadResource(hMod, hResource));
		if(pData) {
			// skip unwanted strings
			for(stringToLoad %= 16; stringToLoad; --stringToLoad) {
				pData += *pData + 1;
			}
			stringLengthToCopy = min(*pData, bufferSize - 1);
			#ifdef UNICODE
				CopyMemory(buffer, pData + 1, stringLengthToCopy * sizeof(WCHAR));
			#else
				CPINFO info = {0};
				GetCPInfo(CP_ACP, &info);
				WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, pData + 1, stringLengthToCopy, buffer, bufferSize, reinterpret_cast<LPCSTR>(info.DefaultChar), NULL);
			#endif
		}
	}
	return stringLengthToCopy;
}


HBITMAP CopyBitmap(HBITMAP hSourceBitmap, BOOL allowNullHandle/* = FALSE*/)
{
	if(!hSourceBitmap && !allowNullHandle) {
		return NULL;
	}

	CDC sourceDC;
	sourceDC.CreateCompatibleDC(NULL);
	CBitmap sourceBitmap;
	HBITMAP hPreviousBitmap1 = NULL;
	SIZE bitmapSize = {0};
	if(hSourceBitmap) {
		sourceBitmap.Attach(hSourceBitmap);
		hPreviousBitmap1 = sourceDC.SelectBitmap(sourceBitmap);
		sourceBitmap.GetSize(bitmapSize);
	}

	CDC targetDC;
	targetDC.CreateCompatibleDC(NULL);
	HDC hScreenDC = GetDC(HWND_DESKTOP);
	HBITMAP hTargetBitmap = CreateCompatibleBitmap(hScreenDC, bitmapSize.cx, bitmapSize.cy);
	ReleaseDC(HWND_DESKTOP, hScreenDC);
	HBITMAP hPreviousBitmap2 = targetDC.SelectBitmap(hTargetBitmap);

	targetDC.BitBlt(0, 0, bitmapSize.cx, bitmapSize.cy, sourceDC, 0, 0, SRCCOPY);

	if(hSourceBitmap) {
		sourceDC.SelectBitmap(hPreviousBitmap1);
		sourceBitmap.Detach();
	}
	targetDC.SelectBitmap(hPreviousBitmap2);

	return hTargetBitmap;
}

HBITMAP CreateBitmapFromDIB(LPBITMAPINFO pBMPInfo)
{
	ATLASSERT_POINTER(pBMPInfo, BITMAPINFO);

	UINT numberOfColors = pBMPInfo->bmiHeader.biClrUsed ? pBMPInfo->bmiHeader.biClrUsed : 1 << pBMPInfo->bmiHeader.biBitCount;
	BOOL numberOfColorsIsGarbage = FALSE;
	if((pBMPInfo->bmiHeader.biClrUsed == 0) && (pBMPInfo->bmiHeader.biBitCount >= 32)) {
		// UINT is too small for such things and changing it to a larger type causes trouble, too
		numberOfColorsIsGarbage = TRUE;
	}

	// calculate the position of the color palette and the DIB bits
	LPVOID pDIBBits = NULL;
	LPRGBQUAD pColorTable = reinterpret_cast<LPRGBQUAD>(reinterpret_cast<LPSTR>(pBMPInfo) + pBMPInfo->bmiHeader.biSize);
	if(pBMPInfo->bmiHeader.biBitCount > 8) {
		pDIBBits = reinterpret_cast<LPVOID>(reinterpret_cast<LPDWORD>(pColorTable + pBMPInfo->bmiHeader.biClrUsed) + ((pBMPInfo->bmiHeader.biCompression == BI_BITFIELDS) ? 3 : 0));
	} else {
		pDIBBits = reinterpret_cast<LPVOID>(pColorTable + numberOfColors);
	}

	// create the palette
	CPalette palette;
	if(!numberOfColorsIsGarbage && (numberOfColors <= 256)) {
		UINT paletteSize = sizeof(LOGPALETTE) + (sizeof(PALETTEENTRY) * numberOfColors);
		LPLOGPALETTE pLogPalette = reinterpret_cast<LPLOGPALETTE>(new BYTE[paletteSize]);
		pLogPalette->palVersion = static_cast<WORD>(0x300);
		pLogPalette->palNumEntries = static_cast<WORD>(numberOfColors);

		for(UINT i = 0; i < numberOfColors; ++i) {
			pLogPalette->palPalEntry[i].peRed = pColorTable[i].rgbRed;
			pLogPalette->palPalEntry[i].peGreen = pColorTable[i].rgbGreen;
			pLogPalette->palPalEntry[i].peBlue = pColorTable[i].rgbBlue;
			pLogPalette->palPalEntry[i].peFlags = 0;
		}

		palette.CreatePalette(pLogPalette);
		delete[] pLogPalette;
	}

	CDC dummyDC;
	dummyDC.CreateCompatibleDC(NULL);

	// apply the palette
	HPALETTE hPreviousPalette = NULL;
	if(!palette.IsNull()) {
		hPreviousPalette = dummyDC.SelectPalette(palette, FALSE);
		dummyDC.RealizePalette();
	}

	HDC hScreenDC = GetDC(HWND_DESKTOP);
	HBITMAP hBitmap = CreateDIBitmap(hScreenDC, &pBMPInfo->bmiHeader, CBM_INIT, pDIBBits, pBMPInfo, DIB_RGB_COLORS);
	ReleaseDC(HWND_DESKTOP, hScreenDC);

	// clean up
	if(hPreviousPalette) {
		dummyDC.SelectPalette(hPreviousPalette, FALSE);
		dummyDC.RealizePalette();
	}
	return hBitmap;
}

HGLOBAL CreateDIBFromBitmap(HBITMAP hBitmap, HPALETTE hPalette, BOOL allow32BPP/* = FALSE*/)
{
	BITMAP bmp = {0};
	GetObject(hBitmap, sizeof(BITMAP), &bmp);

	// fill the header
	BITMAPINFOHEADER BMPHeader = {0};
	BMPHeader.biSize = sizeof(BITMAPINFOHEADER);
	BMPHeader.biWidth = bmp.bmWidth;
	BMPHeader.biHeight = bmp.bmHeight;
	BMPHeader.biPlanes = 1;
	BMPHeader.biCompression = BI_RGB;
	BMPHeader.biBitCount = static_cast<WORD>(bmp.bmPlanes) * static_cast<WORD>(bmp.bmBitsPixel);
	// ensure the bit count is valid
	if(BMPHeader.biBitCount <= 1) {
		BMPHeader.biBitCount = 1;
	} else if(BMPHeader.biBitCount <= 4) {
		BMPHeader.biBitCount = 4;
	} else if(BMPHeader.biBitCount <= 8) {
		BMPHeader.biBitCount = 8;
	} else if(BMPHeader.biBitCount <= 16) {
		BMPHeader.biBitCount = 16;
	} else if(BMPHeader.biBitCount <= 24) {
		BMPHeader.biBitCount = 24;
	} else {
		BMPHeader.biBitCount = (allow32BPP ? 32 : 24);
	}

	// DIBs must be DWORD-aligned
	WORD w = static_cast<WORD>(BMPHeader.biWidth) * BMPHeader.biBitCount;
	DWORD sizeOfBits = static_cast<DWORD>((unsigned) ((w + 31) & (~31)) / 8) * static_cast<DWORD>(BMPHeader.biHeight);

	DWORD sizeOfPalette = 0;
	if(BMPHeader.biBitCount <= 8) {
		#ifdef _WIN64
			sizeOfPalette = (1i64 << BMPHeader.biBitCount) * sizeof(RGBQUAD);
		#else
			sizeOfPalette = (1 << BMPHeader.biBitCount) * sizeof(RGBQUAD);
		#endif
	} else if((BMPHeader.biBitCount == 16) || (BMPHeader.biBitCount == 32)) {
		sizeOfPalette = (3 * sizeof(DWORD));
		BMPHeader.biCompression = BI_BITFIELDS;
	}

	// finally allocate the memory we'll store the data in
	HGLOBAL hGlobalData = GlobalAlloc(GPTR, BMPHeader.biSize + sizeOfPalette + sizeOfBits);
	if(hGlobalData) {
		LPVOID pData = GlobalLock(hGlobalData);
		if(pData) {
			// simply copy the header
			*reinterpret_cast<LPBITMAPINFOHEADER>(pData) = BMPHeader;

			CDC dummyDC;
			dummyDC.CreateCompatibleDC(NULL);

			// apply the palette
			HPALETTE hPreviousPalette = NULL;
			if(hPalette) {
				hPreviousPalette = dummyDC.SelectPalette(hPalette, FALSE);
				dummyDC.RealizePalette();
			}

			// get the bits
			GetDIBits(dummyDC, hBitmap, 0, static_cast<WORD>(BMPHeader.biHeight), reinterpret_cast<LPVOID>(reinterpret_cast<LPBYTE>(pData) + BMPHeader.biSize + sizeOfPalette), reinterpret_cast<LPBITMAPINFO>(pData), DIB_RGB_COLORS);

			// clean up
			if(hPreviousPalette) {
				dummyDC.SelectPalette(hPreviousPalette, FALSE);
				dummyDC.RealizePalette();
			}

			GlobalUnlock(hGlobalData);
		}
	}

	return hGlobalData;
}

HGLOBAL CreateDIBV5FromBitmap(HBITMAP hBitmap, HPALETTE hPalette)
{
	// we create a normal DIB and convert it
	HGLOBAL hDIB = CreateDIBFromBitmap(hBitmap, hPalette, TRUE);
	if(!hDIB) {
		return NULL;
	}

	LPBITMAPINFOHEADER pDIB = static_cast<LPBITMAPINFOHEADER>(GlobalLock(hDIB));
	if(!pDIB) {
		return NULL;
	}

	// calculate the size of the bitmap bits
	ULONG sizeOfBits = ((unsigned) ((pDIB->biWidth * pDIB->biBitCount + 31) & (~31)) / 8) * abs(pDIB->biHeight);

	// calculate the size of the color table
	ULONG sizeOfColorTable = 0;
	ULONG sizeOfColorTableV5 = 0;
	BOOL useBitMasks = FALSE;
	if(pDIB->biCompression == BI_BITFIELDS) {
		if((pDIB->biBitCount == 16) || (pDIB->biBitCount == 32)) {
			// color bit masks for RGB
			sizeOfColorTable = (3 * sizeof(DWORD));
			// bit fields data is a part of BITMAPV5HEADER
			sizeOfColorTableV5 = 0;
			useBitMasks = TRUE;
		} else {
			// this shouldn't happen
			sizeOfColorTable = sizeOfColorTableV5 = 0;
		}
	} else if(pDIB->biCompression == BI_RGB) {
		if(pDIB->biClrUsed) {
			sizeOfColorTable = sizeOfColorTableV5 = pDIB->biClrUsed * sizeof(DWORD);
		} else {
			if(pDIB->biBitCount <= 8) {
				#ifdef _WIN64
					sizeOfColorTable = sizeOfColorTableV5 = (1i64 << pDIB->biBitCount) * sizeof(RGBQUAD);
				#else
					sizeOfColorTable = sizeOfColorTableV5 = (1 << pDIB->biBitCount) * sizeof(RGBQUAD);
				#endif
			} else {
				sizeOfColorTable = sizeOfColorTableV5 = 0;
			}
		}
	} else if(pDIB->biCompression == BI_RLE4) {
		sizeOfColorTable = sizeOfColorTableV5 = 16 * sizeof(DWORD);
	} else if(pDIB->biCompression == BI_RLE8) {
		sizeOfColorTable = sizeOfColorTableV5 = 256 * sizeof(DWORD);
	} else {
		sizeOfColorTable = sizeOfColorTableV5 = 0;
	}

	// finally allocate the memory we'll store the data in
	HGLOBAL hGlobalData = GlobalAlloc(GPTR, sizeof(BITMAPV5HEADER) + sizeOfColorTableV5 + sizeOfBits);
	if(hGlobalData) {
		LPVOID pData = GlobalLock(hGlobalData);
		if(pData) {
			// simply copy the header
			*reinterpret_cast<LPBITMAPINFOHEADER>(pData) = *pDIB;

			// adjust the header size, color space and rendering mode
			reinterpret_cast<LPBITMAPV5HEADER>(pData)->bV5Size = sizeof(BITMAPV5HEADER);
			reinterpret_cast<LPBITMAPV5HEADER>(pData)->bV5CSType = LCS_sRGB;
			reinterpret_cast<LPBITMAPV5HEADER>(pData)->bV5Intent = LCS_GM_IMAGES;

			// adjust the color space information
			if(useBitMasks) {
				reinterpret_cast<LPBITMAPV5HEADER>(pData)->bV5RedMask = *reinterpret_cast<LPDWORD>(&reinterpret_cast<LPBITMAPINFO>(pDIB)->bmiColors[0]);
				reinterpret_cast<LPBITMAPV5HEADER>(pData)->bV5GreenMask = *reinterpret_cast<LPDWORD>(&reinterpret_cast<LPBITMAPINFO>(pDIB)->bmiColors[1]);
				reinterpret_cast<LPBITMAPV5HEADER>(pData)->bV5BlueMask = *reinterpret_cast<LPDWORD>(&reinterpret_cast<LPBITMAPINFO>(pDIB)->bmiColors[2]);
			} else if(sizeOfColorTableV5) {
				// copy the color table
				CopyMemory(reinterpret_cast<LPBYTE>(pData) + sizeof(BITMAPV5HEADER), reinterpret_cast<LPBYTE>(pDIB) + sizeof(BITMAPINFOHEADER), sizeOfColorTableV5);
			}

			// copy the bitmap bits
			CopyMemory(reinterpret_cast<LPBYTE>(pData) + sizeof(BITMAPV5HEADER) + sizeOfColorTableV5, reinterpret_cast<LPBYTE>(pDIB) + sizeof(BITMAPINFOHEADER) + sizeOfColorTable, sizeOfBits);

			GlobalUnlock(hGlobalData);
		}
	}

	// clean up
	GlobalUnlock(hDIB);
	GlobalFree(hDIB);

	return hGlobalData;
}

HPALETTE GetPaletteFromDataObject(IDataObject* pDataObject)
{
	ATLASSUME(pDataObject);
	if(!pDataObject) {
		return NULL;
	}

	FORMATETC format = {CF_PALETTE, NULL, DVASPECT_CONTENT, -1, TYMED_GDI};
	HPALETTE hPalette = NULL;
	STGMEDIUM storageMedium = {0};
	if(pDataObject->GetData(&format, &storageMedium) == S_OK) {
		if(storageMedium.tymed & TYMED_GDI) {
			hPalette = reinterpret_cast<HPALETTE>(storageMedium.hBitmap);
		}
	}

	ReleaseStgMedium(&storageMedium);
	return hPalette;
}

HIMAGELIST SetupStateImageList(HIMAGELIST hStateImageList)
{
	BOOL usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	CTheme themingEngine;
	if(usingThemes) {
		themingEngine.OpenThemeData(NULL, VSCLASS_BUTTON);
	}

	CImageList imageList = hStateImageList;
	SIZE iconSize;
	imageList.GetIconSize(iconSize);

	if(themingEngine.IsThemeNull()) {
		CIcon icon = imageList.ExtractIcon(1);
		if(!icon.IsNull()) {
			ICONINFO iconInfo;
			icon.GetIconInfo(&iconInfo);

			if(iconInfo.hbmColor) {
				CDCHandle compatibleDC = GetDC(HWND_DESKTOP);
				if(compatibleDC) {
					CDC memoryDC;
					memoryDC.CreateCompatibleDC(compatibleDC);
					if(memoryDC) {
						HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(iconInfo.hbmColor);

						for(int y = 0; y < iconSize.cy; ++y) {
							for(int x = 0; x < iconSize.cx; ++x) {
								if(memoryDC.GetPixel(x, y) == RGB(255, 255, 255)) {
									memoryDC.SetPixelV(x, y, RGB(192, 192, 192));
								}
							}
						}

						memoryDC.SelectBitmap(hPreviousBitmap);

						imageList.Add(iconInfo.hbmColor, iconInfo.hbmMask);
					}
					ReleaseDC(HWND_DESKTOP, compatibleDC.Detach());
				}
			}

			if(iconInfo.hbmColor) {
				DeleteObject(iconInfo.hbmColor);
			}
			if(iconInfo.hbmMask) {
				DeleteObject(iconInfo.hbmMask);
			}
			icon.DestroyIcon();
		}
	} else {
		CDCHandle compatibleDC = GetDC(HWND_DESKTOP);
		if(compatibleDC) {
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(compatibleDC);
			if(memoryDC) {
				CBitmap bitmap;
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = iconSize.cx;
				bitmapInfo.bmiHeader.biHeight = -iconSize.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pPixel;
				bitmap.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pPixel), NULL, 0);
				if(bitmap) {
					ZeroMemory(pPixel, iconSize.cx * iconSize.cy * sizeof(RGBQUAD));

					HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(bitmap);

					WTL::CRect rc(0, 0, iconSize.cx, iconSize.cy);
					SIZE partSize;
					themingEngine.GetThemePartSize(memoryDC, BP_CHECKBOX, CBS_MIXEDNORMAL, NULL, TS_TRUE, &partSize);
					rc.OffsetRect((iconSize.cx - partSize.cx) / 2, (iconSize.cy - partSize.cy) / 2);
					themingEngine.DrawThemeBackground(memoryDC, BP_CHECKBOX, CBS_MIXEDNORMAL, &rc);

					memoryDC.SelectBitmap(hPreviousBitmap);

					imageList.Add(bitmap);
				}
			}
			ReleaseDC(HWND_DESKTOP, compatibleDC.Detach());
		}
	}

	if(!themingEngine.IsThemeNull()) {
		themingEngine.CloseThemeData();
	}
	return imageList.Detach();
}

HRESULT IDataObject_GetDropDescription(LPDATAOBJECT pDataObject, DROPDESCRIPTION& dropDescription)
{
	if(!pDataObject) {
		return E_POINTER;
	}

	HRESULT hr = DV_E_FORMATETC;

	FORMATETC format = {0};
	STGMEDIUM medium = {0};

	format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_DROPDESCRIPTION));
	format.dwAspect = DVASPECT_CONTENT;
	format.lindex = -1;
	format.tymed = TYMED_HGLOBAL;
	if(SUCCEEDED(pDataObject->GetData(&format, &medium))) {
		if(medium.hGlobal) {
			DROPDESCRIPTION* pDropDescription = static_cast<DROPDESCRIPTION*>(GlobalLock(medium.hGlobal));
			if(pDropDescription) {
				dropDescription = *pDropDescription;
				hr = S_OK;
			}
			GlobalUnlock(medium.hGlobal);
		}
		ReleaseStgMedium(&medium);
	}
	return hr;
}

HRESULT InvalidateDragWindow(LPDATAOBJECT pDataObject)
{
	if(!pDataObject) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	FORMATETC format = {0};
	STGMEDIUM medium = {0};

	format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
	format.dwAspect = DVASPECT_CONTENT;
	format.lindex = -1;
	format.tymed = TYMED_HGLOBAL;
	if(SUCCEEDED(pDataObject->GetData(&format, &medium))) {
		if(medium.hGlobal) {
			#define DDWM_UPDATEWINDOW		WM_USER + 3     // (wParam = 0, lParam = 0)

			HWND hWndDragWindow = *static_cast<HWND*>(GlobalLock(medium.hGlobal));
			GlobalUnlock(medium.hGlobal);
			if(IsWindow(hWndDragWindow)) {
				PostMessage(hWndDragWindow, DDWM_UPDATEWINDOW, 0, 0);
				hr = S_OK;
			}
		}
		ReleaseStgMedium(&medium);
	}
	return hr;
}

UINT GetDragImageMessage(void)
{
	static UINT message = 0;

	if(message == 0) {
		message = RegisterWindowMessage(DI_GETDRAGIMAGE);
	}
	return message;
}