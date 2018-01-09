//////////////////////////////////////////////////////////////////////
/// \class CWindowEx
/// \author Timo "TimoSoft" Kunze
/// \brief <em>An extension to \c ATL::CWindow</em>
///
/// This class extends \c ATL::CWindow by a \c IsMouseWithin and a \c InternalSetRedraw method.
//////////////////////////////////////////////////////////////////////


#pragma once

#include "stdafx.h"

class CWindowEx :
    public CWindow
{
public:
	CWindowEx(HWND hWnd = NULL) :
	    CWindow(hWnd)
	{
		//
	}

	/// \brief <em>Activates or deactivates automatic redrawing</em>
	///
	/// Activates or deactivates automatic redrawing. The window will be notified that this call was
	/// made from within the program by ourselves.
	///
	/// \param[in] redraw If \c TRUE, the window will redraw itself automatically; otherwise not.
	void InternalSetRedraw(BOOL redraw = TRUE);

	/// \brief <em>Retrieves whether the cursor is located above the window</em>
	///
	/// Retrieves whether the current cursor position lies within the window's bounding rectangle.
	///
	/// \return \c TRUE if the cursor lies within the window rectangle; otherwise \c FALSE.
	BOOL IsMouseWithin(void);
};     // CWindowEx