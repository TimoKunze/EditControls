//////////////////////////////////////////////////////////////////////
/// \class CommonProperties
/// \author Timo "TimoSoft" Kunze
/// \brief <em>The controls' "Common" property page</em>
///
/// This class contains the controls' "Common" property page.
///
/// \sa HotKeyBox, IPAddressBox, TextBox, UpDownTextBox, StringProperties
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


class ATL_NO_VTABLE CommonProperties :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CommonProperties, &CLSID_CommonProperties>,
    public IPropertyPage2Impl<CommonProperties>,
    public CDialogImpl<CommonProperties>
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	CommonProperties();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		enum { IDD = IDD_COMMONPROPERTIES };

		DECLARE_REGISTRY_RESOURCEID(IDR_COMMONPROPERTIES)

		BEGIN_COM_MAP(CommonProperties)
			COM_INTERFACE_ENTRY(IPropertyPage)
		END_COM_MAP()

		BEGIN_MSG_MAP(CommonProperties)
			NOTIFY_CODE_HANDLER(LVN_GETINFOTIP, OnListViewGetInfoTipNotification)
			NOTIFY_CODE_HANDLER(LVN_ITEMCHANGED, OnListViewItemChangedNotification)

			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOA, OnToolTipGetDispInfoNotificationA)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOW, OnToolTipGetDispInfoNotificationW)

			COMMAND_ID_HANDLER(ID_LOADSETTINGS, OnLoadSettingsFromFile)
			COMMAND_ID_HANDLER(ID_SAVESETTINGS, OnSaveSettingsToFile)

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

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c LVN_GETINFOTIP handler</em>
	///
	/// Will be called if a listview control requests an item's tooltip text. We use this handler to display
	/// tooltips for the \c disabledEventsList items.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774835.aspx">LVN_GETINFOTIP</a>
	LRESULT OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c LVN_ITEMCHANGED handler</em>
	///
	/// Will be called if some of the attributes of the specified listview item have been changed. We use
	/// this handler to prevent our custom indeterminate state image from being selected manually.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774845.aspx">LVN_ITEMCHANGED</a>
	LRESULT OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
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

	/// \brief <em>Retrieves a checkmark state that represents a bit's state among multiple values</em>
	///
	/// \param[in] arraysize The number of elements in the array defined by the \c pArray parameter.
	/// \param[in] pArray An array of the values for which to check the bits defined by \c bitsToCheckFor.
	/// \param[in] bitsToCheckFor The bits to check.
	///
	/// \return The masked state image index as required by the listview. The state image symbolizes
	///         &ldquo;unchecked&rdquo;, if none of the values in the array has the specified bits set;
	///         &ldquo;checked&rdquo;, if all values in the array have the specified bits set;
	///         &ldquo;indeterminated&rdquo; otherwise.
	///
	/// \sa SetBit
	static int CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor);
	/// \brief <em>Clears or sets the specified bit within the specified value depending on the specified checkmark state</em>
	///
	/// \param[in] stateImageMask The masked state image index as returned by the listview. If the state
	///            image symbolizes &ldquo;checked&rdquo;, the bit is set; if it symbolizes
	///            &ldquo;unchecked&rdquo;, the bit is cleared; otherwise it remains unchanged.
	/// \param[in,out] value The value in which the bit is modified.
	/// \param[in] bitToSet The bit to set or clear.
	///
	/// \sa CalculateStateImageMask
	static void SetBit(int stateImageMask, LONG& value, LONG bitToSet);

	/// \brief <em>Holds wrappers for the dialog's GUI elements</em>
	struct Controls
	{
		/// \brief <em>Wraps the listview displaying the values for the \c DisabledEvents property</em>
		CCheckListViewCtrl disabledEventsList;
		/// \brief <em>Wraps the listview displaying the values for the \c InvalidKeyCombinations property</em>
		CCheckListViewCtrl invalidKeyCombinationsList;
		/// \brief <em>Wraps the listview displaying the values for the \c DefaultModifierKeys property</em>
		CCheckListViewCtrl defaultModifierKeysList;
		/// \brief <em>Wraps the edit control displaying the \c CurrentValue property's setting</em>
		CEdit currentValueEdit;
		/// \brief <em>Wraps the up down control used to change the \c CurrentValue property's setting</em>
		CUpDownCtrl currentValueUpDown;
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
		/// \brief <em>Specifies the listview control currently being setup in \c LoadSettings</em>
		///
		/// \sa LoadSettings, OnListViewItemChangedNotification
		HWND hWndCheckMarksAreSetFor;

		/// \brief <em>Holds the item indices of the disabled events listbox</em>
		struct DisabledEventsItemIndices
		{
			/// \brief <em>Holds the zero-based index of the item representing the \c deMouseEvents flag</em>
			///
			/// \if UNICODE
			///   \sa EditCtlsLibU::DisabledEventsConstants
			/// \else
			///   \sa EditCtlsLibA::DisabledEventsConstants
			/// \endif
			int deMouseEvents;
			/// \brief <em>Holds the zero-based index of the item representing the \c deClickEvents flag</em>
			///
			/// \if UNICODE
			///   \sa EditCtlsLibU::DisabledEventsConstants
			/// \else
			///   \sa EditCtlsLibA::DisabledEventsConstants
			/// \endif
			int deClickEvents;
			/// \brief <em>Holds the zero-based index of the item representing the \c deKeyboardEvents flag</em>
			///
			/// \if UNICODE
			///   \sa EditCtlsLibU::DisabledEventsConstants
			/// \else
			///   \sa EditCtlsLibA::DisabledEventsConstants
			/// \endif
			int deKeyboardEvents;
			/// \brief <em>Holds the zero-based index of the item representing the \c deBeforeDrawText flag</em>
			///
			/// \if UNICODE
			///   \sa EditCtlsLibU::DisabledEventsConstants
			/// \else
			///   \sa EditCtlsLibA::DisabledEventsConstants
			/// \endif
			int deBeforeDrawText;
			/// \brief <em>Holds the zero-based index of the item representing the \c deScrolling flag</em>
			///
			/// \if UNICODE
			///   \sa EditCtlsLibU::DisabledEventsConstants
			/// \else
			///   \sa EditCtlsLibA::DisabledEventsConstants
			/// \endif
			int deScrolling;
			/// \brief <em>Holds the zero-based index of the item representing the \c deTextChangedEvents flag</em>
			///
			/// \if UNICODE
			///   \sa EditCtlsLibU::DisabledEventsConstants
			/// \else
			///   \sa EditCtlsLibA::DisabledEventsConstants
			/// \endif
			int deTextChangedEvents;
			/// \brief <em>Holds the zero-based index of the item representing the \c deValueChangingEvents flag</em>
			///
			/// \if UNICODE
			///   \sa EditCtlsLibU::DisabledEventsConstants
			/// \else
			///   \sa EditCtlsLibA::DisabledEventsConstants
			/// \endif
			int deValueChangingEvents;

			DisabledEventsItemIndices()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				deMouseEvents = -2;
				deClickEvents = -2;
				deKeyboardEvents = -2;
				deBeforeDrawText = -2;
				deScrolling = -2;
				deTextChangedEvents = -2;
				deValueChangingEvents = -2;
			}
		} disabledEventsItemIndices;

		Properties()
		{
			hWndCheckMarksAreSetFor = NULL;
		}
	} properties;

private:
};     // CommonProperties

OBJECT_ENTRY_AUTO(CLSID_CommonProperties, CommonProperties)