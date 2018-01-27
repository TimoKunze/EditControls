// CWindowEx2.cpp: an extension to WTL::CWindowEx

#include "stdafx.h"
#include "CWindowEx2.h"


void CWindowEx2::InternalSetRedraw(BOOL redraw/* = TRUE*/)
{
	SendMessage(WM_SETREDRAW, redraw, 71216);
}

BOOL CWindowEx2::IsMouseWithin(void)
{
	RECT windowRectangle;
	GetClientRect(&windowRectangle);
	POINT mousePosition;
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);
	// check whether the mouse is within the rect
	return PtInRect(&windowRectangle, mousePosition);
}