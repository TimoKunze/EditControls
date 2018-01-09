//////////////////////////////////////////////////////////////////////
/// \class ICreditsProvider
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between about dialog and control</em>
///
/// This interface allows the about dialog to ask the control for the credits and other infos that
/// it'll display.
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa AboutDlg, HotKeyBox, IPAddressBox, TextBox, UpDownTextBox
//////////////////////////////////////////////////////////////////////


#pragma once


class ICreditsProvider
{
public:
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	virtual CAtlString GetAuthors(void) = 0;
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	virtual CAtlString GetHomepage(void) = 0;
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	virtual CAtlString GetPaypalLink(void) = 0;
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	virtual CAtlString GetSpecialThanks(void) = 0;
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	virtual CAtlString GetThanks(void) = 0;
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	virtual CAtlString GetVersion(void) = 0;
};     // ICreditsProvider