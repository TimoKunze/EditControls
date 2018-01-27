//////////////////////////////////////////////////////////////////////
/// \class CWindowEx2
/// \author Timo "TimoSoft" Kunze
/// \brief <em>An extension to \c WTL::CWindowEx</em>
///
/// This class extends \c WTL::CWindowEx by a \c IsMouseWithin and a \c InternalSetRedraw method.
//////////////////////////////////////////////////////////////////////


#pragma once

#include "stdafx.h"

class CWindowEx2 :
    public CWindowEx
{
public:
	CWindowEx2(HWND hWnd = NULL) :
	    CWindowEx(hWnd)
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
};     // CWindowEx2