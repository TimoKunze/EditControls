//////////////////////////////////////////////////////////////////////
/// \class StringProperties
/// \author Timo "TimoSoft" Kunze
/// \brief <em>The controls' "String Properties" property page</em>
///
/// This class contains the controls' "String Properties" property page.
///
/// \sa HotKeyBox, IPAddressBox, TextBox, UpDownTextBox, CommonProperties
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "EditCtlsU.h"
#else
	#include "EditCtlsA.h"
#endif
#include "HotKeyBox.h"
#include "IPAddressBox.h"
#include "TextBox.h"
#include "UpDownTextBox.h"


class ATL_NO_VTABLE StringProperties :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<StringProperties, &CLSID_StringProperties>,
    public IPropertyPage2Impl<StringProperties>,
    public CDialogImpl<StringProperties>
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	StringProperties();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		enum { IDD = IDD_STRINGPROPERTIES };

		DECLARE_REGISTRY_RESOURCEID(IDR_STRINGPROPERTIES)

		BEGIN_COM_MAP(StringProperties)
			COM_INTERFACE_ENTRY(IPropertyPage)
			//COM_INTERFACE_ENTRY2(IPropertyPage2, IPropertyPage)
		END_COM_MAP()

		BEGIN_MSG_MAP(StringProperties)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOA, OnToolTipGetDispInfoNotificationA)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOW, OnToolTipGetDispInfoNotificationW)

			COMMAND_ID_HANDLER(ID_LOADSETTINGS, OnLoadSettingsFromFile)
			COMMAND_ID_HANDLER(ID_SAVESETTINGS, OnSaveSettingsToFile)

			COMMAND_HANDLER(IDC_PROPERTYCOMBO, CBN_SELCHANGE, OnPropertySelChange)
			COMMAND_CODE_HANDLER(EN_CHANGE, OnChange)
			REFLECT_NOTIFICATIONS()
		END_MSG_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPropertyPage
	///
	//@{
	/// \brief <em>Creates the property page's dialog box</em>
	///
	/// Creates the property page's dialog box window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetObjects,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682250.aspx">IPropertyPage::Activate</a>
	virtual HRESULT STDMETHODCALLTYPE Activate(HWND hWndParent, LPCRECT pRect, BOOL modal);
	/// \brief <em>Applies the current values to the controls</em>
	///
	/// Applies the current values to the underlying objects associated with the property page.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetObjects,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms691284.aspx">IPropertyPage::Apply</a>
	virtual HRESULT STDMETHODCALLTYPE Apply(void);
	/// \brief <em>Provides the controls this property page is associated to</em>
	///
	/// \param[in] objects The number of objects defined by the \c ppControls array.
	/// \param[in] ppControls An array of \c IUnknown pointers each identifying an associated control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Activate, Apply,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms678529.aspx">IPropertyPage::SetObjects</a>
	virtual HRESULT STDMETHODCALLTYPE SetObjects(ULONG objects, IUnknown** ppControls);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPropertyPage2
	///
	//@{
	/// \brief <em>Specifies which field is to receive the focus when the property page is activated</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms690353.aspx">IPropertyPage2::EditProperty</a>
	virtual HRESULT STDMETHODCALLTYPE EditProperty(DISPID dispID);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c TTN_GETDISPINFOA handler</em>
	///
	/// Will be called if a tooltip control requests the text to display. We use this handler to display
	/// tooltips for the toolbar buttons.
	///
	/// \sa OnToolTipGetDispInfoNotificationW,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOW handler</em>
	///
	/// Will be called if a tooltip control requests the text to display. We use this handler to display
	/// tooltips for the toolbar buttons.
	///
	/// \sa OnToolTipGetDispInfoNotificationA,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c ID_LOADSETTINGS handler</em>
	///
	/// Will be called if the \c ID_LOADSETTINGS command was initiated.
	/// We use this handler to execute the command.
	LRESULT OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c ID_SAVESETTINGS handler</em>
	///
	/// Will be called if the \c ID_SAVESETTINGS command was initiated.
	/// We use this handler to execute the command.
	LRESULT OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c CBN_SELCHANGE handler</em>
	///
	/// Will be called if the property combo box control's selection was changed.
	/// We use this handler to update the value displayed in the property edit control.
	LRESULT OnPropertySelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c EN_CHANGE handler</em>
	///
	/// Will be called if an edit control's content was changed.
	/// We use this handler to set the dirty flag.
	LRESULT OnChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies all values to all associated controls</em>
	///
	/// \sa LoadSettings
	void ApplySettings(void);
	/// \brief <em>Loads the current properties' settings from the associated controls</em>
	///
	/// \sa ApplySettings
	void LoadSettings(void);

	/// \brief <em>Holds wrappers for the dialog's GUI elements</em>
	struct Controls
	{
		/// \brief <em>Wraps the combo box control used to select the string property to edit</em>
		CComboBox propertyCombo;
		/// \brief <em>Wraps the edit control displaying a string property's setting</em>
		CEdit propertyEdit;
		/// \brief <em>Wraps the toolbar control providing the buttons to save/load settings</em>
		CToolBarCtrl toolbar;
		/// \brief <em>Holds the \c toolbar control's enabled icons</em>
		CImageList imagelistEnabled;
		/// \brief <em>Holds the \c toolbar control's disabled icons</em>
		CImageList imagelistDisabled;

		~Controls()
		{
			if(!imagelistEnabled.IsNull()) {
				imagelistEnabled.Destroy();
			}
			if(!imagelistDisabled.IsNull()) {
				imagelistDisabled.Destroy();
			}
		}
	} controls;

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Specifies the string property to pre-select in the property combo box</em>
		DISPID propertyToEdit;
		/// \brief <em>The item ID of the currently selected item in the property combo box</em>
		int selectedPropertyItemID;
		/// \brief <em>The buffered value of the \c IIPAddressBox::Address property</em>
		CAtlString address;
		/// \brief <em>The buffered value of the \c CueBanner properties</em>
		CAtlString cueBanner;
		/// \brief <em>The buffered value of the \c Text properties</em>
		CAtlString text;

		Properties()
		{
			propertyToEdit = -1;
			selectedPropertyItemID = -1;
			address = TEXT("");
			cueBanner = TEXT("");
			text = TEXT("");
		}
	} properties;

private:
};     // StringProperties

OBJECT_ENTRY_AUTO(CLSID_StringProperties, StringProperties)