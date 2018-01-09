//////////////////////////////////////////////////////////////////////
/// \class HyperLinkEx
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A derivative of \c WTL::CHyperLink that allows custom tooltips</em>
///
/// \sa AboutDlg
//////////////////////////////////////////////////////////////////////


#pragma once

#include "stdafx.h"


class HyperLinkEx :
    public CHyperLink
{
public:
	/// \brief <em>Retrieves the control's current tooltip text</em>
	///
	/// \param[in] pBuffer A buffer that receives the current tooltip text.
	/// \param[in] bufferlength The length of the buffer specified by \c pBuffer.
	///
	/// \return \c True on success; otherwise \c false.
	///
	/// \sa SetToolTipText
	bool GetToolTipText(LPTSTR pBuffer, int bufferlength) const
	{
		#ifndef _WIN32_WCE
			if(m_tip.IsWindow()) {
				TCHAR tmp[2048];
				m_tip.GetText(tmp, *this, 1);
				#if _SECURE_ATL
					ATL::Checked::tcscpy_s(pBuffer, bufferlength, tmp);
				#else
					lstrcpyn(pBuffer, tmp, bufferlength);
				#endif
				return true;
			}
		#endif // !_WIN32_WCE
		return false;
	}

	/// \brief <em>Sets the control's tooltip text</em>
	///
	/// \param[in] pToolTipText The tooltip text to set.
	///
	/// \return \c True on success; otherwise \c false.
	///
	/// \sa GetToolTipText
	bool SetToolTipText(LPCTSTR pToolTipText)
	{
		if(!m_lpstrLabel) {
			CalcLabelRect();
		}
		#ifndef _WIN32_WCE
			if(m_tip.IsWindow()) {
				m_tip.Activate(TRUE);
				TTTOOLINFO toolInfo = {0};
				#ifdef UNICODE
					toolInfo.cbSize = TTTOOLINFOW_V2_SIZE;
				#else
					toolInfo.cbSize = TTTOOLINFOA_V2_SIZE;
				#endif
				toolInfo.uFlags = 0;
				toolInfo.hwnd = *this;
				toolInfo.uId = 1;
				toolInfo.rect = m_rcLink;
				toolInfo.lpszText = const_cast<LPTSTR>(pToolTipText);
				toolInfo.hinst = ModuleHelper::GetResourceInstance();
				m_tip.AddTool(&toolInfo);
				return true;
			}
		#endif // !_WIN32_WCE
		return false;
	}

private:
};     // HyperLinkEx