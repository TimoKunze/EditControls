//////////////////////////////////////////////////////////////////////
/// \file stdafx.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Standard includes</em>
///
/// This file contains the standard includes. It also contains basic defines like \c WINVER and
/// the defines for controlling debugging mechanisms like memleak detection.
///
/// \todo Improve documentation ("See also" sections).
//////////////////////////////////////////////////////////////////////


#pragma once
#pragma warning(disable: 4127)     // conditional expression is constant
#pragma warning(disable: 4505)     // unreferenced local function has been removed
#pragma warning(disable: 4512)     // assignment operator could not be generated
#pragma warning(disable: 6031)     // return value ignored
#pragma warning(disable: 6400)     // using lstrcmpi to perform a case-insensitive compare to constant string yields unexpected results in non-English locales
#pragma warning(disable: 6401)     // using CompareString to perform a case-insensitive compare to constant string yields unexpected results in non-English locales
#pragma warning(push)
#pragma warning(disable: 4995)     // name was marked as #pragma deprecated
#pragma warning(disable: 6011)     // dereferencing NULL pointer
#pragma warning(disable: 6387)     // argument might be 0

#ifndef STRICT
	/// \brief <em>Disables some implicit data type conversions and other stuff</em>
	///
	/// Defining \c STRICT advises the compiler to not allow passing an \c HWND where an \c HDC
	/// argument is required and similar things.
	#define STRICT
#endif

/// \brief <em>Controls compatibility</em>
///
/// Defining \c NTDDI_VERSION as \c NTDDI_WIN10 activates parts of the Windows API that require Windows 10
/// or lower.
#define NTDDI_VERSION NTDDI_WIN10

#include <winsdkver.h>

/// \brief <em>Tells the compiler we're using apartment threading</em>
#define _ATL_APARTMENT_THREADED

/// \brief <em>Makes certain \c CString constructors explicit, preventing any unintentional conversions</em>
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS

/// \brief <em>???</em>
// TODO: Write some docu.
#define OEMRESOURCE

#ifdef _DEBUG
	/// \brief <em>Enables interface leak detection</em>
	#define _ATL_DEBUG_REFCOUNT
	#include <windows.h>
	/// \brief <em>Enables mem leak detection</em>
	#include <atldbgmem.h>
#else
	#define USE_STL
#endif

#ifdef USE_STL
	// STL headers
	#include <vector>
	#include <unordered_map>
	#include <algorithm>
#endif

// ATL headers
#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <atltypes.h>
#include <atlhost.h>
#include <atlctl.h>
#include <atlstr.h>
#include <atlsafe.h>     // CComSafeArray
#ifndef USE_STL
	#include <atlcoll.h>     // ATL collection classes
#endif

#include <strsafe.h>

#include <commoncontrols.h>     // IImageList
#include <wincodec.h>     // WIC, used for Aero drag images

// WTL headers
#include <atlapp.h>     // message loop, interfaces, general app stuff
#include <atlctrls.h>     // standard and common control classes
#include <atlctrlx.h>     // hyperlink label, bitmap button, check list view, and other controls
#include <atldlgs.h>     // common dialogs
#include <atlgdi.h>     // DC classes, GDI object classes
#include <atltheme.h>     // Windows XP theme classes
#include <atlmisc.h>     // WTL ports of CPoint, CRect, CSize, CString, etc.
//#include <atlexcept.h>     // all we need for exceptions

// OLE headers
#include <olectl.h>

// shell headers
#include <shlobj.h>
#pragma warning(pop)