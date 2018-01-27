// AboutDlg.cpp: the controls' about dialog

#include "stdafx.h"
#include "AboutDlg.h"


void AboutDlg::SetOwner(ICreditsProvider* pOwner)
{
	properties.pOwner = pOwner;
}


LRESULT AboutDlg::OnCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	HWND hWnd = reinterpret_cast<HWND>(lParam);
	if(controls.IsStatic(hWnd)) {
		// bolden the font and select it into the dc
		HDC hDC = reinterpret_cast<HDC>(wParam);
		LOGFONT font = {0};
		// use the current font as base
		GetObject(CWindow(hWnd).GetFont(), sizeof(font), &font);
		font.lfWeight = FW_BOLD;
		HFONT hFont = CreateFontIndirect(&font);
		if(hFont) {
			SelectObject(hDC, hFont);
			DeleteObject(hFont);
		}

		// set the color to COLOR_HIGHLIGHT
		SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHT));
		SetBkMode(hDC, TRANSPARENT);

		// make the background transparent
		return reinterpret_cast<LRESULT>(GetStockObject(HOLLOW_BRUSH));
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT AboutDlg::OnInitDialog(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// attach to the controls
	controls.author = GetDlgItem(IDC_AUTHOR);
	controls.gitHubRepository.SubclassWindow(GetDlgItem(IDC_GITHUB));
	controls.homepage.SubclassWindow(GetDlgItem(IDC_HOMEPAGE));
	controls.more = GetDlgItem(IDC_MORE);
	controls.name = GetDlgItem(IDC_NAME);
	controls.paypal.SubclassWindow(GetDlgItem(IDC_PAYPAL));
	controls.thanks = GetDlgItem(IDC_THANKS);
	controls.specialThanks = GetDlgItem(IDC_SPECIALTHANKS);

	// setup the images
	CStatic(GetDlgItem(IDC_ABOUT)).SetBitmap(images.hAboutImage);
	CStatic(GetDlgItem(IDC_PAYPAL)).SetBitmap(images.hPaypalImage);

	if(properties.pOwner) {
		// get the credits and fill the controls
		SetDlgItemText(IDC_VERSION, properties.pOwner->GetVersion());
		SetDlgItemText(IDC_AUTHORS, properties.pOwner->GetAuthors());
		SetDlgItemText(IDC_THANKSTO, properties.pOwner->GetThanks());
		SetDlgItemText(IDC_SPECIALTHANKSTO, properties.pOwner->GetSpecialThanks());

		// tell the control the link
		CAtlString str = properties.pOwner->GetHomepage();
		controls.homepage.SetHyperLink(str);
		// make it look pretty
		str.Replace(TEXT("https://"), TEXT(""));
		controls.homepage.SetLabel(str);
		// also set a tooltip text
		str = TEXT("Browse to ") + str + TEXT("...");
		controls.homepage.SetToolTipText((LPCTSTR) str);

		controls.gitHubRepository.SetHyperLink(TEXT("https://github.com/TimoKunze/EditControls"));
		controls.gitHubRepository.SetLabel(TEXT("Fork me on GitHub"));
		controls.gitHubRepository.SetToolTipText(TEXT("Browse the GitHub repository..."));

		controls.paypal.SetHyperLink(properties.pOwner->GetPaypalLink());
		controls.paypal.SetToolTipText(TEXT("Make a donation to the author of this software."));
	}

	return TRUE;
}


HRESULT AboutDlg::OnOK(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	// the user clicked Ok, so close this dialog
	EndDialog(0);
	return S_OK;
}