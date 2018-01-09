// CWindowEx.cpp: an extension to ATL::CWindow

#include "stdafx.h"
#include "CWindowEx.h"


void CWindowEx::InternalSetRedraw(BOOL redraw/* = TRUE*/)
{
	SendMessage(WM_SETREDRAW, redraw, 71216);
}

BOOL CWindowEx::IsMouseWithin(void)
{
	RECT windowRectangle;
	GetClientRect(&windowRectangle);
	POINT mousePosition;
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);
	// check whether the mouse is within the rect
	return PtInRect(&windowRectangle, mousePosition);
}