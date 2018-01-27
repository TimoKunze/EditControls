//////////////////////////////////////////////////////////////////////
/// \class AboutDlg
/// \author Timo "TimoSoft" Kunze
/// \brief <em>The controls' about dialog</em>
///
/// This class contains the controls' about dialog, which contains credits, a link to our website,
/// a Paypal link and other things.
///
/// \todo Improve documentation.
///
/// \sa ICreditsProvider, HotKeyBox, IPAddressBox, TextBox, UpDownTextBox
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#include "helpers.h"
#include "HyperLinkEx.h"
#include "ICreditsProvider.h"


class AboutDlg :
    public CDialogImpl<AboutDlg>
{
public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		enum { IDD = IDD_ABOUTBOX };

		BEGIN_MSG_MAP(AboutDlg)
			MESSAGE_HANDLER(WM_CTLCOLORSTATIC, OnCtlColorStatic)
			MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

			COMMAND_ID_HANDLER(IDOK, OnOK)

			REFLECT_NOTIFICATIONS_EX()
		END_MSG_MAP()
	#endif

	/// \brief <em>Sets the dialog's owner</em>
	///
	/// The owner provides the data needed to fill the dialog.
	///
	/// \param[in] pOwner The \c ICreditsProvider object that owns this dialog.
	void SetOwner(ICreditsProvider* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CTLCOLORSTATIC handler</em>
	///
	/// Will be called if a static control is about to be drawn. Used to set the font and forecolor for
	/// the headline labels.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms651165.aspx">WM_CTLCOLORSTATIC</a>
	LRESULT OnCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_INITDIALOG handler</em>
	///
	/// Will be called if a dialog's content needs to be initialized. Used to initialize the controls with
	/// the data provided by this dialog's owner.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms645428.aspx">WM_INITDIALOG</a>
	LRESULT OnInitDialog(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>Ok button click-handler</em>
	///
	/// Will be called if the user clicks the Ok button.
	HRESULT OnOK(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds wrappers for the dialog's GUI elements</em>
	struct Controls
	{
		/// \brief <em>Wraps the static label displaying the caption for the "Written by" section</em>
		CStatic author;
		/// \brief <em>Wraps the static label displaying the caption for the "Further information" section</em>
		CStatic more;
		/// \brief <em>Wraps the static label displaying the caption for the intro section</em>
		CStatic name;
		/// \brief <em>Wraps the link label displaying the Paypal link</em>
		HyperLinkEx paypal;
		/// \brief <em>Wraps the static label displaying the caption for the "Thanks go to" section</em>
		CStatic thanks;
		/// \brief <em>Wraps the static label displaying the caption for the "Special thanks go to" section</em>
		CStatic specialThanks;
		/// \brief <em>Wraps the link label displaying the URL of our website</em>
		HyperLinkEx homepage;
		/// \brief <em>Wraps the link label displaying the URL of the GitHub repository</em>
		HyperLinkEx gitHubRepository;

		/// \brief <em>Retrieves whether a window handle belongs to one of the dialog's static controls</em>
		///
		/// \param[in] hWnd The window handle to check.
		///
		/// \return \c TRUE if the window handle identifies one of the dialog's static labels.
		///
		/// \sa OnCtlColorStatic
		BOOL IsStatic(HWND hWnd)
		{
			if(hWnd == author) {
				return TRUE;
			}
			if(hWnd == more) {
				return TRUE;
			}
			if(hWnd == name) {
				return TRUE;
			}
			if(hWnd == thanks) {
				return TRUE;
			}
			if(hWnd == specialThanks) {
				return TRUE;
			}

			return FALSE;
		}
	} controls;

	/// \brief <em>Holds the dialog's images</em>
	struct Images
	{
		/// \brief <em>Holds the about image</em>
		HBITMAP hAboutImage;
		/// \brief <em>Holds the Paypal image</em>
		HBITMAP hPaypalImage;

		Images()
		{
			hAboutImage = LoadJPGResource(IDB_ABOUT);
			hPaypalImage = LoadJPGResource(IDB_PAYPAL);
		}

		~Images()
		{
			if(hAboutImage) {
				DeleteObject(static_cast<HGDIOBJ>(hAboutImage));
			}
			if(hPaypalImage) {
				DeleteObject(static_cast<HGDIOBJ>(hPaypalImage));
			}
		}
	} images;

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The dialog's owner</em>
		ICreditsProvider* pOwner;

		Properties()
		{
			pOwner = NULL;
		}
	} properties;
};     // AboutDlg