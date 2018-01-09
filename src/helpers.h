//////////////////////////////////////////////////////////////////////
/// \file helpers.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Global helper functions, macros and other stuff</em>
///
/// This file contains global helper functions, macros, types, constants and other stuff.
///
/// \todo Rewrite \c ConvertIntegerToString() so that the returned string is freed automatically.
/// \todo \c GetButtonStateBitField(): Do swapped mouse buttons get handled correctly?
/// \todo \c GetButtonStateBitField(): According to MSDN \c MK_ALT and \c MK_RBUTTON are never set.
/// \todo \c GetResStringWithNumber(): Allow more parameters to insert.
/// \todo Improve documentation ("See also" sections).
/// \todo Move callback functions to the appropriate class (as static methods).
//////////////////////////////////////////////////////////////////////


#pragma once

#include <comutil.h>
#include "DispIDs.h"

class HotKeyBox;
class IPAddressBox;
class TextBox;
class UpDownTextBox;


//////////////////////////////////////////////////////////////////////
/// \name Extended ATL message map macros
///
//@{
#define INDEXED_MESSAGE_HANDLER(msg, func) \
	if(uMsg == msg) { \
		bHandled = TRUE; \
		lResult = func(dwMsgMapID, uMsg, wParam, lParam, bHandled); \
		if(bHandled) { \
			return TRUE; \
		} \
	}
//@}
//////////////////////////////////////////////////////////////////////


#ifdef NDEBUG
	#define ATLASSERT_POINTER(pointer, type)
	#define ATLASSERT_ARRAYPOINTER(pointer, type, count)
#else
	/// \brief <em>Ensures a given pointer is valid</em>
	///
	/// Throws an assertion if \c pointer is \c NULL or we don't have read access for this address.
	///
	/// \param[in] pointer The pointer to check.
	/// \param[in] type The type of the data that \c pointer points to.
	#define ATLASSERT_POINTER(pointer, type) \
	    ATLASSERT(((pointer) != NULL) && AtlIsValidAddress(reinterpret_cast<LPCVOID>(pointer), sizeof(type), FALSE))
	/// \brief <em>Ensures a given pointer is valid</em>
	///
	/// Throws an assertion if \c pointer is \c NULL or we don't have read access for this address.
	///
	/// \param[in] pointer The pointer to check.
	/// \param[in] type The type of the data that \c pointer points to.
	/// \param[in] count The number of elements in the array pointed to by \c pointer.
	#define ATLASSERT_ARRAYPOINTER(pointer, type, count) \
	    ATLASSERT(((pointer) != NULL) && AtlIsValidAddress(reinterpret_cast<LPCVOID>(pointer), count * sizeof(type), FALSE))
#endif

/// \brief <em>Notifies the container that a property is about to change</em>
///
/// Notifies the container that a property is about to change. The container can cancel the change.
/// This is used with data-binding. It allows the data source to check whether the specified property
/// actually can be edited before the change is applied.
///
/// \param[in] dispid The property to check.
///
/// \remarks If the edit is cancelled, the calling method is left with \c S_FALSE.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/e7h956et.aspx">CComControl::FireOnRequestEdit</a>
#define PUTPROPPROLOG(dispid) if(FireOnRequestEdit(dispid) == S_FALSE) return S_FALSE

/// \brief <em>Converts from \c BOOL to \c VARIANT_BOOL</em>
///
/// \param[in] value The \c BOOL value to convert.
///
/// \return \c VARIANT_TRUE if \c value is \c TRUE; otherwise \c VARIANT_FALSE.
///
/// \sa VARIANTBOOL2BOOL
#define BOOL2VARIANTBOOL(value) \
    ((value) ? VARIANT_TRUE : VARIANT_FALSE)

/// \brief <em>Converts from \c VARIANT_BOOL to \c BOOL</em>
///
/// \param[in] value The \c VARIANT_BOOL value to convert.
///
/// \return \c TRUE if \c value is \c VARIANT_TRUE; otherwise \c FALSE.
///
/// \sa BOOL2VARIANTBOOL
#define VARIANTBOOL2BOOL(value) \
    ((value) != VARIANT_FALSE)

/// \brief <em>Frees memory and sets the pointer to \c NULL</em>
///
/// Frees the memory \c pMem points to and sets \c pMem to \c NULL avoiding illegal pointers.
///
/// \param[in] pMem Pointer to the memory to free. It is set to \c NULL.
#define SECUREFREE(pMem) \
    free(pMem); \
    pMem = NULL

/// \brief <em>Calculates the size (in bytes) of a string</em>
///
/// Calculates the correct size (in bytes) of \c str considering the different width of an Unicode
/// and an ANSI character.
///
/// \param[in] str: The string for which to calculate the size.
///
/// \return The size of \c str in bytes.
#define SECURESIZEOFSTRING(str) \
    sizeof(str) / sizeof(TCHAR)

/// \brief <em>Calculates the number of elements in an array</em>
///
/// \param[in] array The array for which to calculate the number of elements.
/// \param[in] type The data type of the array's elments.
///
/// \return The number of elements in \c array.
#define COUNTELEMENTSOFARRAY(array, type) \
    (sizeof array) / sizeof(type)

/// \brief <em>Generates VB6 compatible bit fields for the state of the mouse and shift buttons</em>
///
/// \c BitField holds the state of the three (left, right, middle) mouse buttons and the three
/// (Shift, Ctrl, Alt) shift buttons. This function splits those information into two separate bit
/// fields - one for the mouse buttons and one for the shift buttons.
///
/// \param[in] bitField A bit field holding all button states.
/// \param[out] mouseButtons A bit field that will be used to hold the state of the mouse buttons.
/// \param[out] shiftButtons A bit field that will be used to hold the state of the shift buttons.
///
/// \remarks This function usually is used if \c bitField was generated by OLE code (e. g. in OLE
///          drag'n'drop code).\n
///          The two resulting bit fields are structured in the style of Visual Basic 6. This means:
///          - Bit 1 holds the state of the Shift button/left mouse button
///          - Bit 2 holds the state of the Ctrl button/right mouse button
///          - Bit 3 holds the state of the Alt button/middle mouse button
///
/// \sa WPARAM2BUTTONANDSHIFT
void OLEKEYSTATE2BUTTONANDSHIFT(int bitField, SHORT& mouseButtons, SHORT& shiftButtons);

/// \brief <em>Generates VB6 compatible bit fields for the state of the mouse and shift buttons</em>
///
/// \c BitField holds the state of the five (left, right, middle, extended1, extended2) mouse buttons and
/// the three (Shift, Ctrl, Alt) shift buttons. This function splits those information into two separate
/// bit fields - one for the mouse buttons and one for the shift buttons.
///
/// \param[in] bitField A bit field holding all button states. If set to -1, the buttons' states
///            will be retrieved by calling the \c GetAsyncKeyState API function.
/// \param[out] mouseButtons A bit field that will be used to hold the state of the mouse buttons.
/// \param[out] shiftButtons A bit field that will be used to hold the state of the shift buttons.
///
/// \remarks This function usually is used if \c bitField is a parameter of a windows message
///          (e. g. \c WM_MOUSEMOVE).\n
///          The two resulting bit fields are structured in the style of Visual Basic 6. This means:
///          - Bit 1 holds the state of the Shift button/left mouse button
///          - Bit 2 holds the state of the Ctrl button/right mouse button
///          - Bit 3 holds the state of the Alt button/middle mouse button
///          - Bit 6 holds the state of the first extended mouse button
///          - Bit 7 holds the state of the second extended mouse button
///
/// \sa OLEKEYSTATE2BUTTONANDSHIFT
void WPARAM2BUTTONANDSHIFT(int bitField, SHORT& mouseButtons, SHORT& shiftButtons);

/// \brief <em>Holds DLL version information</em>
typedef struct DllVersion
{
	/// \brief Specifies the platform the DLL was built for.
	enum Platform {
		/// \brief The DLL was not built for any special Windows platform.
		Windows,
		/// \brief The DLL was built specifically for Windows NT.
		WindowsNT,
		/// \brief The DLL's version information didn't contain any platform information.
		Unknown
	} targetPlatform;
	/// Specifies the DLL's major version number.
	DWORD majorVersion;
	/// Specifies the DLL's minor version number.
	DWORD minorVersion;
	/// Specifies the DLL's build number.
	DWORD buildNumber;

	DllVersion()
	{
		// 'majorVersion == -1' means 'we don't contain valid data'
		majorVersion = static_cast<DWORD>(-1);
	}

	/// \brief <em>Retrieves whether the contained data is valid</em>
	///
	/// \return \c TRUE if the structure contains valid data; otherwise \c FALSE.
	BOOL IsValid(void)
	{
		return (majorVersion != static_cast<DWORD>(-1));
	}
} DllVersion;


//////////////////////////////////////////////////////////////////////
/// \name Conversions
///
//@{
/// \brief <em>Converts an integer value to a string</em>
///
/// \param[in] value The number to convert.
///
/// \return A pointer to the string containing the number specified by \c value. The calling function is
///         responsible for freeing this string.
LPTSTR ConvertIntegerToString(char value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(int value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(long value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(__int64 value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned char value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned int value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned long value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned __int64 value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \param[in] str The string to convert.
/// \param[out] value The converted integer.
void ConvertStringToInteger(LPTSTR str, char& value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \overload
void ConvertStringToInteger(LPTSTR str, int& value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \overload
void ConvertStringToInteger(LPTSTR str, long& value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \overload
void ConvertStringToInteger(LPTSTR str, __int64& value);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \param[in] hDC The device context to use for conversion.
/// \param[in] pixels The value to convert in pixels.
/// \param[in] vertical If set to \c TRUE, \c pixels specifies a value along the y-axis; otherwise
///            it specifies a value along the x-axis.
///
/// \return The converted value in twips.
long ConvertPixelsToTwips(HDC hDC, long pixels, BOOL vertical);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
POINT ConvertPixelsToTwips(HDC hDC, POINT &pixels);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
RECT ConvertPixelsToTwips(HDC hDC, RECT &pixels);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
SIZE ConvertPixelsToTwips(HDC hDC, SIZE &pixels);
/// \brief <em>Converts twip coordinates to pixel coordinates</em>
///
/// \param[in] hDC The device context to use for conversion.
/// \param[in] twips The value to convert in twips.
/// \param[in] vertical If set to \c TRUE, \c twips specifies a value along the y-axis; otherwise
///            it specifies a value along the x-axis.
///
/// \return The converted value in pixels.
long ConvertTwipsToPixels(HDC hDC, long twips, BOOL vertical);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
POINT ConvertTwipsToPixels(HDC hDC, POINT &twips);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
RECT ConvertTwipsToPixels(HDC hDC, RECT &twips);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
SIZE ConvertTwipsToPixels(HDC hDC, SIZE &twips);
/// \brief <em>Translates an \c OLE_COLOR type color into a \c COLORREF type color</em>
///
/// \param[in] oleColor The \c OLE_COLOR type color to translate.
/// \param[in] hPalette The color palette to use.
///
/// \return The translated \c COLORREF type color.
COLORREF OLECOLOR2COLORREF(OLE_COLOR oleColor, HPALETTE hPalette = NULL);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Error throwing
///
//@{
/// \brief <em>Throws a COM error</em>
///
/// Throws a COM error that can be handled by Visual Basic's \c Err object.
///
/// \param[in] hError The COM error code.
/// \param[in] source The object throwing the error.
/// \param[in] title A string that might be used as title for the error message.
/// \param[in] description The error description.
/// \param[in] helpContext The context in the help file associated with the error.
/// \param[in] helpFileName The help file associated with the error.
///
/// \return An \c HRESULT error code.
///
/// \remarks If \c description is set to \c NULL, the error description is retrieved from the system.
HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, LPTSTR description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, UINT description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, LPTSTR description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, UINT description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
///
/// \param[in] errorCode The Win32 error code.
/// \param[in] source The object throwing the error.
/// \param[in] title A string that might be used as title for the error message.
/// \param[in] helpContext The context in the help file associated with the error.
/// \param[in] helpFileName The help file associated with the error.
HRESULT DispatchError(DWORD errorCode, REFCLSID source, LPTSTR title, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(DWORD errorCode, REFCLSID source, UINT title, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
//@}
//////////////////////////////////////////////////////////////////////

/// \brief <em>Reads a \c VARIANT from a stream</em>
///
/// Reads a \c VARIANT from a stream. The only data types supported are \c VT_BOOL, \c VT_DATE, \c VT_I2
/// and \c VT_I4. The data must be stored in a space-efficient way.
///
/// \param[in] pStream The stream to read from.
/// \param[in] expectedVarType The expected data type. The only types supported are \c VT_BOOL, \c VT_DATE,
///            \c VT_I2 and \c VT_I4.
/// \param[out] pVariant Receives the read data.
///
/// \return An \c HRESULT error code.
///
/// \sa ReadVariantFromStream_Legacy, WriteVariantToStream
HRESULT ReadVariantFromStream(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
/// \brief <em>Reads a \c VARIANT from a stream</em>
///
/// Reads a \c VARIANT from a stream. The only data types supported are \c VT_BOOL, \c VT_DATE, \c VT_I2
/// and \c VT_I4. The data must be stored with 16 Byte per \c VARIANT.
///
/// \param[in] pStream The stream to read from.
/// \param[in] expectedVarType The expected data type. The only types supported are \c VT_BOOL, \c VT_DATE,
///            \c VT_I2 and \c VT_I4.
/// \param[out] pVariant Receives the read data.
///
/// \return An \c HRESULT error code.
///
/// \sa ReadVariantFromStream, WriteVariantToStream
HRESULT ReadVariantFromStream_Legacy(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
/// \brief <em>Writes a \c VARIANT to a stream</em>
///
/// Writes a \c VARIANT to a stream. The only data types supported are \c VT_BOOL, \c VT_DATE, \c VT_I2
/// and \c VT_I4. The data is stored in a space-efficient way.
///
/// \param[in] pStream The stream to write to.
/// \param[in] pVariant The data to write.
///
/// \return An \c HRESULT error code.
///
/// \sa ReadVariantFromStream, ReadVariantFromStream_Legacy
HRESULT WriteVariantToStream(__in LPSTREAM pStream, __in LPVARIANT pVariant);
/// \brief <em>Retrieves the shift and mouse button states</em>
///
/// Retrieves a bit field holding the states of the five mouse (left, right, middle, x1, x2) and the
/// three shift buttons (Shift, Ctrl, Alt).
///
/// \return A bit field holding the buttons' states.
WPARAM GetButtonStateBitField(void);
/// \brief <em>Retrieves a DLL's version</em>
///
/// Retrieves detailed version information about a given DLL.
///
/// \param[in] dllName The DLL's name.
///
/// \return The version information.
DllVersion GetDllVersion(LPCTSTR dllName);
/// \brief <em>Loads a string resource</em>
///
/// Loads a given string resource and converts it into a COM string.
///
/// \param[in] stringToLoad The resource ID of the string to load.
///
/// \return The loaded string.
CComBSTR GetResString(UINT stringToLoad);
/// \brief <em>Inserts an integer number into a string</em>
///
/// Loads a given string resource, inserts a given integer number and returns the string converted
/// into a COM string. The insertion point must be tagged by '%i'.
///
/// \param[in] stringToLoad The resource ID of the string to load.
/// \param[in] numberToInsert The number to insert.
///
/// \return The loaded string containing the number.
CComBSTR GetResStringWithNumber(UINT stringToLoad, int numberToInsert);
/// \brief <em>Loads the specified JPG resource as bitmap</em>
///
/// \param[in] jpgToLoad The resource ID of the JPG to load.
///
/// \return The loaded JPG as bitmap.
HBITMAP LoadJPGResource(UINT jpgToLoad);
/// \brief <em>Retrieves whether an accelerator index is a valid one</em>
///
/// \param[in] acceleratorIndex The accelerator index to check.
/// \param[in] pUpDownTextBox The \c UpDownTextBox object containing the accelerator.
///
/// \return \c TRUE if the accelerator index is valid; otherwise \c FALSE.
BOOL IsValidAcceleratorIndex(int acceleratorIndex, UpDownTextBox* pUpDownTextBox);
/// \brief <em>Retrieves whether an accelerator index is a valid one</em>
///
/// \overload
///
/// \param[in] acceleratorIndex The accelerator index to check.
/// \param[in] hWndUpDown The up down control containing the accelerator.
///
/// \return \c TRUE if the accelerator index is valid; otherwise \c FALSE.
BOOL IsValidAcceleratorIndex(int acceleratorIndex, HWND hWndUpDown);
/// \brief <em>Loads the specified string resource from comctl32.dll paying attention to localization</em>
///
/// \param[in] stringToLoad The resource ID of the string to load.
/// \param[in,out] buffer A buffer of size \c bufferSize that the string will be copied to.
/// \param[in] bufferSize The number of characters that the buffer specified by \c buffer may receive.
///
/// \return The number of characters that were copied to the buffer.
///
/// \sa HotKeyBox::get_Text
int LoadLocalizedStringResource(UINT stringToLoad, LPTSTR buffer, int bufferSize);

/// \brief <em>Retrieves a copy of a given bitmap</em>
///
/// \param[in] hSourceBitmap The bitmap to copy.
/// \param[in] allowNullHandle If \c TRUE, the function creates a bitmap even if \c hSourceBitmap is
///            set to \c NULL; otherwise it will simply return \c NULL, if \c hSourceBitmap is \c NULL.
///
/// \return The copied bitmap.
///
/// \sa TargetOLEDataObject
HBITMAP CopyBitmap(HBITMAP hSourceBitmap, BOOL allowNullHandle = FALSE);
/// \brief <em>Creates a bitmap from a \c BITMAPINFO struct</em>
///
/// \param[in] pBMPInfo The \c BITMAPINFO struct to create the bitmap from.
///
/// \return The created bitmap.
///
/// \remarks The DIB bits must follow the \c BITMAPINFO struct.
///
/// \sa TargetOLEDataObject::GetData, CreateDIBFromBitmap, CreateDIBV5FromBitmap,
///     <a href="https://msdn.microsoft.com/en-us/library/ms532284.aspx">BITMAPINFO</a>
HBITMAP CreateBitmapFromDIB(__in LPBITMAPINFO pBMPInfo);
/// \brief <em>Creates a DIB from a bitmap</em>
///
/// Returns a global memory object containing a \c BITMAPINFO struct that defines a DIB which was
/// created from a given DDB.
///
/// \param[in] hBitmap The DDB to create the DIB from.
/// \param[in] hPalette The palette to use.
/// \param[in] allow32BPP If \c TRUE, the maximum color depth will be 32 bits per pixel; otherwise it
///            will be 24 bits per pixel.
///
/// \return A global memory object containing the DIB data.
///
/// \sa SourceOLEDataObject::SetData, TargetOLEDataObject::SetData, CreateBitmapFromDIB,
///     CreateDIBV5FromBitmap,
///     <a href="https://msdn.microsoft.com/en-us/library/ms532284.aspx">BITMAPINFO</a>
HGLOBAL CreateDIBFromBitmap(HBITMAP hBitmap, HPALETTE hPalette, BOOL allow32BPP = FALSE);
/// \brief <em>Creates a DIB (version 5) from a bitmap</em>
///
/// Returns a global memory object containing a \c BITMAPINFO struct that defines a DIB which was
/// created from a given DDB. This function differs from \c CreateDIBFromBitmap in that it uses a
/// \c BITMAPV5HEADER header instead of \c BITMAPINFOHEADER.
///
/// \param[in] hBitmap The DDB to create the DIB from.
/// \param[in] hPalette The palette to use.
///
/// \return A global memory object containing the DIB data.
///
/// \sa SourceOLEDataObject::SetData, TargetOLEDataObject::SetData, CreateBitmapFromDIB,
///     CreateDIBFromBitmap,
///     <a href="https://msdn.microsoft.com/en-us/library/ms532284.aspx">BITMAPINFO</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/ms532331.aspx">BITMAPV5HEADER</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/ms532290.aspx">BITMAPINFOHEADER</a>
HGLOBAL CreateDIBV5FromBitmap(HBITMAP hBitmap, HPALETTE hPalette);
/// \brief <em>Retrieves \c CF_PALETTE data from an \c IDataObject implementation</em>
///
/// Queries the specified \c IDataObject implementation for data in the \c CF_PALETTE format.
///
/// \param[in] pDataObject The \c IDataObject implementation to work on.
///
/// \return The handle of the palette data, if the data object consists such data; otherwise \c NULL.
///
/// \sa TargetOLEDataObject,
///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/ms649013.aspx">CF_PALETTE</a>
HPALETTE GetPaletteFromDataObject(IDataObject* pDataObject);
/// \brief <em>Adds a state image for indeterminated state to the specified image list</em>
///
/// A normal state image list contains images for &ldquo;checked&rdquo; and &ldquo;unchecked&rdquo;
/// states only. This method adds an image for &ldquo;indeterminate&rdquo; state.
///
/// \param[in] hStateImageList The image list to which the image will be added.
///
/// \return The handle of the modified image list (usually the same as \c hStateImageList).
HIMAGELIST SetupStateImageList(HIMAGELIST hStateImageList);
/// \brief <em>Retrieves the \c CFSTR_DROPDESCRIPTION data stored by a data object</em>
///
/// \param[in] pDataObject The data object that contains the data.
/// \param[out] dropDescription The \c CFSTR_DROPDESCRIPTION data.
///
/// \return An \c HRESULT error code.
HRESULT IDataObject_GetDropDescription(__in LPDATAOBJECT pDataObject, DROPDESCRIPTION& dropDescription);
/// \brief <em>Invalidates the window that is displaying the drag image</em>
///
/// \param[in] pDataObject The data object that contains the handle to the drag window in the format
///            \c DragWindow.
///
/// \return An \c HRESULT error code.
HRESULT InvalidateDragWindow(__in LPDATAOBJECT pDataObject);
/// \brief <em>Retrieves the current ID of the \c DI_GETDRAGIMAGE message</em>
///
/// \return The message's current ID.
///
/// \sa TextBox::OnGetDragImage, TextBox::OLEDrag,
///     <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
UINT GetDragImageMessage(void);