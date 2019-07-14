//////////////////////////////////////////////////////////////////////
/// \class UpDownTextBox
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Combines an \c Edit control with a \c msctls_updown32 control</em>
///
/// This class combines an \c Edit control with a \c msctls_updown32 control and makes the result
/// accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
///
/// \if UNICODE
///   \sa EditCtlsLibU::IUpDownTextBox
/// \else
///   \sa EditCtlsLibA::IUpDownTextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "EditCtlsU.h"
#else
	#include "EditCtlsA.h"
#endif
#include "_IUpDownTextBoxEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "CommonProperties.h"
#include "StringProperties.h"
#include "UpDownAccelerator.h"
#include "UpDownAccelerators.h"
#include "VirtualUpDownAccelerator.h"
#include "VirtualUpDownAccelerators.h"
#include "TargetOLEDataObject.h"


class ATL_NO_VTABLE UpDownTextBox : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IUpDownTextBox, &IID_IUpDownTextBox, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IUpDownTextBox, &IID_IUpDownTextBox, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<UpDownTextBox>,
    public IOleControlImpl<UpDownTextBox>,
    public IOleObjectImpl<UpDownTextBox>,
    public IOleInPlaceActiveObjectImpl<UpDownTextBox>,
    public IViewObjectExImpl<UpDownTextBox>,
    public IOleInPlaceObjectWindowlessImpl<UpDownTextBox>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<UpDownTextBox>,
    public Proxy_IUpDownTextBoxEvents<UpDownTextBox>,
    public IPersistStorageImpl<UpDownTextBox>,
    public IPersistPropertyBagImpl<UpDownTextBox>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<UpDownTextBox>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_UpDownTextBox, &__uuidof(_IUpDownTextBoxEvents), &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_UpDownTextBox, &__uuidof(_IUpDownTextBoxEvents), &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<UpDownTextBox>,
    public CComCoClass<UpDownTextBox, &CLSID_UpDownTextBox>,
    public CComCompositeControl<UpDownTextBox>,
    public IPerPropertyBrowsingImpl<UpDownTextBox>,
    public IDropTarget,
    public ICategorizeProperties,
    public ICreditsProvider
{
public:
	/// \brief <em>The contained edit control</em>
	CContainedWindow containedEdit;
	/// \brief <em>The contained up down control</em>
	CContainedWindow containedUpDown;

	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	UpDownTextBox();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		enum { IDD = IDD_UPDWNTXTBOX };

		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_IMEMODE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_UPDOWNTEXTBOX)

		/*#ifdef UNICODE
			DECLARE_WND_CLASS(TEXT("UpDownTextBoxU"))
		#else
			DECLARE_WND_CLASS(TEXT("UpDownTextBoxA"))
		#endif*/

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(UpDownTextBox)
			COM_INTERFACE_ENTRY(IUpDownTextBox)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
		END_COM_MAP()

		BEGIN_PROP_MAP(UpDownTextBox)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AcceptNumbersOnly", DISPID_UPDWNTXTBOX_ACCEPTNUMBERSONLY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AlwaysShowSelection", DISPID_UPDWNTXTBOX_ALWAYSSHOWSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_UPDWNTXTBOX_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutomaticallyCorrectValue", DISPID_UPDWNTXTBOX_AUTOMATICALLYCORRECTVALUE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AutomaticallySetText", DISPID_UPDWNTXTBOX_AUTOMATICALLYSETTEXT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AutoScrolling", DISPID_UPDWNTXTBOX_AUTOSCROLLING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BackColor", DISPID_UPDWNTXTBOX_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("Base", DISPID_UPDWNTXTBOX_BASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_UPDWNTXTBOX_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CancelIMECompositionOnSetFocus", DISPID_UPDWNTXTBOX_CANCELIMECOMPOSITIONONSETFOCUS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CharacterConversion", DISPID_UPDWNTXTBOX_CHARACTERCONVERSION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CompleteIMECompositionOnKillFocus", DISPID_UPDWNTXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("CueBanner", DISPID_UPDWNTXTBOX_CUEBANNER, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("CurrentValue", DISPID_UPDWNTXTBOX_CURRENTVALUE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DetectDoubleClicks", DISPID_UPDWNTXTBOX_DETECTDOUBLECLICKS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisabledBackColor", DISPID_UPDWNTXTBOX_DISABLEDBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_UPDWNTXTBOX_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisabledForeColor", DISPID_UPDWNTXTBOX_DISABLEDFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("DisplayCueBannerOnFocus", DISPID_UPDWNTXTBOX_DISPLAYCUEBANNERONFOCUS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_UPDWNTXTBOX_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DoOEMConversion", DISPID_UPDWNTXTBOX_DOOEMCONVERSION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Enabled", DISPID_UPDWNTXTBOX_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_UPDWNTXTBOX_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_UPDWNTXTBOX_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("GroupDigits", DISPID_UPDWNTXTBOX_GROUPDIGITS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("HAlignment", DISPID_UPDWNTXTBOX_HALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HotTracking", DISPID_UPDWNTXTBOX_HOTTRACKING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("HoverTime", DISPID_UPDWNTXTBOX_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IMEMode", DISPID_UPDWNTXTBOX_IMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("LeftMargin", DISPID_UPDWNTXTBOX_LEFTMARGIN, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Maximum", DISPID_UPDWNTXTBOX_MAXIMUM, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaxTextLength", DISPID_UPDWNTXTBOX_MAXTEXTLENGTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Minimum", DISPID_UPDWNTXTBOX_MINIMUM, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Modified", DISPID_UPDWNTXTBOX_MODIFIED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_UPDWNTXTBOX_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_UPDWNTXTBOX_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Orientation", DISPID_UPDWNTXTBOX_ORIENTATION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessArrowKeys", DISPID_UPDWNTXTBOX_PROCESSARROWKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_UPDWNTXTBOX_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ReadOnlyTextBox", DISPID_UPDWNTXTBOX_READONLYTEXTBOX, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_UPDWNTXTBOX_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightMargin", DISPID_UPDWNTXTBOX_RIGHTMARGIN, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_UPDWNTXTBOX_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_UPDWNTXTBOX_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("Text", DISPID_UPDWNTXTBOX_TEXT, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("UpDownPosition", DISPID_UPDWNTXTBOX_UPDOWNPOSITION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_UPDWNTXTBOX_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("WrapAtBoundaries", DISPID_UPDWNTXTBOX_WRAPATBOUNDARIES, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(UpDownTextBox)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IUpDownTextBoxEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(UpDownTextBox)
			MESSAGE_HANDLER(WM_CTLCOLOREDIT, OnCtlColorEdit)
			MESSAGE_HANDLER(WM_CTLCOLORSTATIC, OnCtlColorStatic)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_HSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
			MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_VSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

			NOTIFY_CODE_HANDLER(UDN_DELTAPOS, OnDeltaPosNotification)

			COMMAND_CODE_HANDLER(EN_ALIGN_LTR_EC, OnAlign)
			COMMAND_CODE_HANDLER(EN_ALIGN_RTL_EC, OnAlign)
			COMMAND_CODE_HANDLER(EN_CHANGE, OnChange)
			COMMAND_CODE_HANDLER(EN_ERRSPACE, OnErrSpace)
			COMMAND_CODE_HANDLER(EN_KILLFOCUS, OnKillFocus)
			COMMAND_CODE_HANDLER(EN_MAXTEXT, OnMaxText)
			COMMAND_CODE_HANDLER(EN_SETFOCUS, OnSetFocus)
			COMMAND_CODE_HANDLER(EN_UPDATE, OnUpdate)

			CHAIN_MSG_MAP(CComCompositeControl<UpDownTextBox>)
			ALT_MSG_MAP(1)
			MESSAGE_HANDLER(WM_CHAR, OnEditChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnEditContextMenu)
			MESSAGE_HANDLER(WM_IME_CHAR, OnEditIMEChar)
			MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnEditInputLangChange)
			MESSAGE_HANDLER(WM_KEYDOWN, OnEditKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnEditKeyUp)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnEditLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnEditLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnEditLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnEditMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnEditMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnEditMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnEditMouseHover)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnEditMouseWheel)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnEditMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnEditMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnEditMouseWheel)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnEditRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnEditRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnEditRButtonUp)
			MESSAGE_HANDLER(WM_SETFOCUS, OnEditSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnEditSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnEditSetRedraw)
			MESSAGE_HANDLER(WM_SETTEXT, OnEditSetText)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnEditKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnEditKeyUp)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnEditXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnEditXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnEditXButtonUp)

			MESSAGE_HANDLER(EM_SETCUEBANNER, OnSetCueBanner)

			ALT_MSG_MAP(2)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnUpDownContextMenu)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnUpDownLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnUpDownLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnUpDownLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnUpDownMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnUpDownMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnUpDownMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnUpDownMouseHover)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnUpDownMouseWheel)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnUpDownMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnUpDownMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnUpDownMouseWheel)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnUpDownRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnUpDownRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnUpDownRButtonUp)
			MESSAGE_HANDLER(WM_SETFOCUS, OnUpDownSetFocus)
			MESSAGE_HANDLER(WM_SETREDRAW, OnUpDownSetRedraw)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnUpDownXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnUpDownXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnUpDownXButtonUp)

			MESSAGE_HANDLER(UDM_SETACCEL, OnSetAccel)
			MESSAGE_HANDLER(UDM_SETRANGE, OnSetRange)
			MESSAGE_HANDLER(UDM_SETRANGE32, OnSetRange)
		END_MSG_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the control should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IUpDownTextBox
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Accelerators property</em>
	///
	/// Retrieves a collection object wrapping the control's accelerators.\n
	/// The accelerators are used if the user changes the control's current value using the up down arrows.
	/// Each accelerator defines a number of seconds, that the user must change the value continuously,
	/// before the accelerator is used, and an increment by which the control's current value is changed on
	/// each step.\n
	/// The accelerators are sorted by ascending activation time and ascending step size.
	///
	/// \param[out] ppAccelerators Receives the collection object's \c IUpDownAccelerators implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa UpDownAccelerators, get_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE get_Accelerators(IUpDownAccelerators** ppAccelerators);
	/// \brief <em>Retrieves the current setting of the \c AcceptNumbersOnly property</em>
	///
	/// Retrieves whether the contained edit control accepts all kind of text or only numbers. If set to
	/// \c VARIANT_TRUE, only numbers, otherwise all text is accepted.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AcceptNumbersOnly, get_Text, get_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE get_AcceptNumbersOnly(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AcceptNumbersOnly property</em>
	///
	/// Sets whether the contained edit control accepts all kind of text or only numbers. If set to
	/// \c VARIANT_TRUE, only numbers, otherwise all text is accepted.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AcceptNumbersOnly, put_Text, put_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE put_AcceptNumbersOnly(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AlwaysShowSelection property</em>
	///
	/// Retrieves whether the selected text will be highlighted even if the contained edit control doesn't
	/// have the focus. If set to \c VARIANT_TRUE, selected text is drawn as selected if the contained edit
	/// control does not have the focus; otherwise it's drawn as normal text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained edit window.
	///
	/// \sa put_AlwaysShowSelection, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_AlwaysShowSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AlwaysShowSelection property</em>
	///
	/// Sets whether the selected text will be highlighted even if the contained edit control doesn't
	/// have the focus. If set to \c VARIANT_TRUE, selected text is drawn as selected if the contained edit
	/// control does not have the focus; otherwise it's drawn as normal text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained edit window.
	///
	/// \sa get_AlwaysShowSelection, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_AlwaysShowSelection(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by
	/// the \c AppearanceConstants enumeration except \c aDefault is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Appearance, get_BorderStyle, EditCtlsLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, EditCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration except \c aDefault is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Appearance, put_BorderStyle, EditCtlsLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, EditCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c AutomaticallyCorrectValue property</em>
	///
	/// Retrieves whether the control ensures, that the value displayed by the contained edit control always
	/// is a valid value. In detail, the edit control's content is validated if the \c Minimum or \c Maximum
	/// property is changed and if the control loses the focus.\n
	/// If set to \c VARIANT_TRUE, the displayed text is validated automatically and changed if necessary;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AutomaticallyCorrectValue, get_Text, get_Minimum, get_Maximum, get_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE get_AutomaticallyCorrectValue(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutomaticallyCorrectValue property</em>
	///
	/// Sets whether the control ensures, that the value displayed by the contained edit control always
	/// is a valid value. In detail, the edit control's content is validated if the \c Minimum or \c Maximum
	/// property is changed and if the control loses the focus.\n
	/// If set to \c VARIANT_TRUE, the displayed text is validated automatically and changed if necessary;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutomaticallyCorrectValue, put_Text, put_Minimum, put_Maximum, put_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE put_AutomaticallyCorrectValue(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AutomaticallySetText property</em>
	///
	/// Retrieves whether the contained edit control's text automatically is set to the contained up down
	/// control's current value, if this value changes. If set to \c VARIANT_TRUE, the text is changed
	/// automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa put_AutomaticallySetText, get_Text, get_CurrentValue, get_GroupDigits
	virtual HRESULT STDMETHODCALLTYPE get_AutomaticallySetText(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutomaticallySetText property</em>
	///
	/// Sets whether the contained edit control's text automatically is set to the contained up down
	/// control's current value, if this value changes. If set to \c VARIANT_TRUE, the text is changed
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa get_AutomaticallySetText, put_Text, put_CurrentValue, put_GroupDigits
	virtual HRESULT STDMETHODCALLTYPE put_AutomaticallySetText(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoScrolling property</em>
	///
	/// Retrieves the directions into which the contained edit control scrolls automatically, if the caret
	/// reaches the borders of the contained edit control's client area. Any combination of the values
	/// defined by the \c AutoScrollingConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained edit window.
	///
	/// \if UNICODE
	///   \sa put_AutoScrolling, EditCtlsLibU::AutoScrollingConstants
	/// \else
	///   \sa put_AutoScrolling, EditCtlsLibA::AutoScrollingConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AutoScrolling(AutoScrollingConstants* pValue);
	/// \brief <em>Sets the \c AutoScrolling property</em>
	///
	/// Sets the directions into which the contained edit control scrolls automatically, if the caret
	/// reaches the borders of the contained edit control's client area. Any combination of the values
	/// defined by the \c AutoScrollingConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained edit window.
	///
	/// \if UNICODE
	///   \sa get_AutoScrolling, EditCtlsLibU::AutoScrollingConstants
	/// \else
	///   \sa get_AutoScrolling, EditCtlsLibA::AutoScrollingConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_AutoScrolling(AutoScrollingConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BackColor, get_ForeColor, get_DisabledBackColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BackColor, put_ForeColor, put_DisabledBackColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c Base property</em>
	///
	/// Retrieves the format in which the current value is displayed in the contained edit control. Any of
	/// the values defined by the \c BaseConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Base, get_CurrentValue, get_AutomaticallySetText, EditCtlsLibU::BaseConstants
	/// \else
	///   \sa put_Base, get_CurrentValue, get_AutomaticallySetText, EditCtlsLibA::BaseConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Base(BaseConstants* pValue);
	/// \brief <em>Sets the \c Base property</em>
	///
	/// Sets the format in which the current value is displayed in the contained edit control. Any of
	/// the values defined by the \c BaseConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Base, put_CurrentValue, put_AutomaticallySetText, EditCtlsLibU::BaseConstants
	/// \else
	///   \sa get_Base, put_CurrentValue, put_AutomaticallySetText, EditCtlsLibA::BaseConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Base(BaseConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values defined
	/// by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, get_Appearance, EditCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, EditCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values defined by
	/// the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, put_Appearance, EditCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, EditCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CancelIMECompositionOnSetFocus property</em>
	///
	/// Retrieves whether the contained edit control cancels the IME composition string when it receives the
	/// focus. If set to \c VARIANT_TRUE, the composition string is canceled; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CancelIMECompositionOnSetFocus, get_IMEMode, get_CompleteIMECompositionOnKillFocus
	virtual HRESULT STDMETHODCALLTYPE get_CancelIMECompositionOnSetFocus(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CancelIMECompositionOnSetFocus property</em>
	///
	/// Sets whether the contained edit control cancels the IME composition string when it receives the
	/// focus. If set to \c VARIANT_TRUE, the composition string is canceled; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CancelIMECompositionOnSetFocus, put_IMEMode, put_CompleteIMECompositionOnKillFocus
	virtual HRESULT STDMETHODCALLTYPE put_CancelIMECompositionOnSetFocus(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CharacterConversion property</em>
	///
	/// Retrieves the kind of conversion that is applied to characters that are typed into the contained
	/// edit control. Any of the values defined by the \c CharacterConversionConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_CharacterConversion, get_DoOEMConversion, get_Text,
	///       EditCtlsLibU::CharacterConversionConstants
	/// \else
	///   \sa put_CharacterConversion, get_DoOEMConversion, get_Text,
	///       EditCtlsLibA::CharacterConversionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CharacterConversion(CharacterConversionConstants* pValue);
	/// \brief <em>Sets the \c CharacterConversion property</em>
	///
	/// Sets the kind of conversion that is applied to characters that are typed into the contained
	/// edit control. Any of the values defined by the \c CharacterConversionConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_CharacterConversion, put_DoOEMConversion, put_Text,
	///       EditCtlsLibU::CharacterConversionConstants
	/// \else
	///   \sa get_CharacterConversion, put_DoOEMConversion, put_Text,
	///       EditCtlsLibA::CharacterConversionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_CharacterConversion(CharacterConversionConstants newValue);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c CompleteIMECompositionOnKillFocus property</em>
	///
	/// Retrieves whether the contained edit control completes the IME composition string when it loses the
	/// focus. If set to \c VARIANT_TRUE, the composition string is completed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CompleteIMECompositionOnKillFocus, get_IMEMode, get_CancelIMECompositionOnSetFocus
	virtual HRESULT STDMETHODCALLTYPE get_CompleteIMECompositionOnKillFocus(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CompleteIMECompositionOnKillFocus property</em>
	///
	/// Sets whether the contained edit control completes the IME composition string when it loses the
	/// focus. If set to \c VARIANT_TRUE, the composition string is completed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CompleteIMECompositionOnKillFocus, put_IMEMode, put_CancelIMECompositionOnSetFocus
	virtual HRESULT STDMETHODCALLTYPE put_CompleteIMECompositionOnKillFocus(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CueBanner property</em>
	///
	/// Retrieves the contained edit control's textual cue.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to an bug in Windows XP and Windows Server 2003, cue banners won't work on those
	///          systems if East Asian language and complex script support is installed.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa put_CueBanner, get_Text, get_DisplayCueBannerOnFocus
	virtual HRESULT STDMETHODCALLTYPE get_CueBanner(BSTR* pValue);
	/// \brief <em>Sets the \c CueBanner property</em>
	///
	/// Sets the contained edit control's textual cue.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to an bug in Windows XP and Windows Server 2003, cue banners won't work on those
	///          systems if East Asian language and complex script support is installed.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_CueBanner, put_Text, put_DisplayCueBannerOnFocus
	virtual HRESULT STDMETHODCALLTYPE put_CueBanner(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c CurrentValue property</em>
	///
	/// Retrieves the control's current value.
	///
	/// \param[out] invalidValue If \c VARIANT_TRUE, the value currently displayed by the contained edit
	///             control is invalid; otherwise not.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa put_CurrentValue, get_Text, get_AutomaticallySetText, get_AcceptNumbersOnly, Raise_ValueChanging,
	///     Raise_ValueChanged
	virtual HRESULT STDMETHODCALLTYPE get_CurrentValue(VARIANT_BOOL* pInvalidValue = NULL, LONG* pValue = NULL);
	/// \brief <em>Sets the \c CurrentValue property</em>
	///
	/// Sets the control's current value.
	///
	/// \param[out] invalidValue Ignored.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa get_CurrentValue, put_Text, put_AutomaticallySetText, put_AcceptNumbersOnly, Raise_ValueChanging,
	///     Raise_ValueChanged
	virtual HRESULT STDMETHODCALLTYPE put_CurrentValue(VARIANT_BOOL* pInvalidValue = NULL, LONG newValue = 0);
	/// \brief <em>Retrieves the current setting of the \c DetectDoubleClicks property</em>
	///
	/// Retrieves whether double clicks are enabled or disabled. If set to \c VARIANT_TRUE, double clicks are
	/// accepted; otherwise all clicks are handled as single clicks.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects the up-down window only.\n
	///          Enabling double-clicks may lead to accidental double-clicks.
	///
	/// \sa put_DetectDoubleClicks, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	virtual HRESULT STDMETHODCALLTYPE get_DetectDoubleClicks(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DetectDoubleClicks property</em>
	///
	/// Enables or disables double clicks. If set to \c VARIANT_TRUE, double clicks are accepted; otherwise
	/// all clicks are handled as single clicks.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects the up-down window only.\n
	///          Enabling double-clicks may lead to accidental double-clicks.
	///
	/// \sa get_DetectDoubleClicks, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	virtual HRESULT STDMETHODCALLTYPE put_DetectDoubleClicks(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledBackColor property</em>
	///
	/// Retrieves the color used as the control's background color, if the control is read-only or
	/// disabled. If set to -1, the system's default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DisabledBackColor, get_Enabled, get_ReadOnlyTextBox, get_DisabledForeColor, get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_DisabledBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c DisabledBackColor property</em>
	///
	/// Sets the color used as the control's background color, if the control is read-only or
	/// disabled. If set to -1, the system's default color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DisabledBackColor, put_Enabled, put_ReadOnlyTextBox, put_DisabledForeColor, put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_DisabledBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, EditCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, EditCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, EditCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, EditCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledForeColor property</em>
	///
	/// Retrieves the color used as the control's text color, if the control is read-only. If set to -1, the
	/// system's default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On current versions of Windows this property has no effect if the control is disabled.
	///
	/// \sa put_DisabledForeColor, get_Enabled, get_ReadOnlyTextBox, get_DisabledBackColor, get_ForeColor
	virtual HRESULT STDMETHODCALLTYPE get_DisabledForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c DisabledForeColor property</em>
	///
	/// Sets the color used as the control's text color, if the control is read-only. If set to -1, the
	/// system's default color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On current versions of Windows this property has no effect if the control is disabled.
	///
	/// \sa get_DisabledForeColor, put_Enabled, put_ReadOnlyTextBox, put_DisabledBackColor, put_ForeColor
	virtual HRESULT STDMETHODCALLTYPE put_DisabledForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayCueBannerOnFocus property</em>
	///
	/// Retrieves whether the contained edit control's textual cue is displayed if the control has the
	/// keyboard focus. If set to \c VARIANT_TRUE, the textual cue is displayed if the control has the
	/// keyboard focus; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to an bug in Windows XP and Windows Server 2003, cue banners won't work on those
	///          systems if East Asian language and complex script support is installed.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa put_DisplayCueBannerOnFocus, get_CueBanner
	virtual HRESULT STDMETHODCALLTYPE get_DisplayCueBannerOnFocus(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayCueBannerOnFocus property</em>
	///
	/// Sets whether the contained edit control's textual cue is displayed if the control has the
	/// keyboard focus. If set to \c VARIANT_TRUE, the textual cue is displayed if the control has the
	/// keyboard focus; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to an bug in Windows XP and Windows Server 2003, cue banners won't work on those
	///          systems if East Asian language and complex script support is installed.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_DisplayCueBannerOnFocus, put_CueBanner
	virtual HRESULT STDMETHODCALLTYPE put_DisplayCueBannerOnFocus(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DontRedraw property</em>
	///
	/// Retrieves whether automatic redrawing of the control is enabled or disabled. If set to
	/// \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE get_DontRedraw(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DontRedraw property</em>
	///
	/// Enables or disables automatic redrawing of the control. Disabling redraw while doing large changes
	/// on the control may increase performance. If set to \c VARIANT_FALSE, the control will redraw itself
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE put_DontRedraw(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DoOEMConversion property</em>
	///
	/// Retrieves whether the contained edit control's text is converted from the Windows character set to
	/// the OEM character set and then back to the Windows character set. Such a conversion ensures proper
	/// character conversion when the application calls the \c CharToOem function to convert a Windows string
	/// in the contained edit control to OEM characters. This property is most useful if the contained edit
	/// control contains file names that will be used on file systems that do not support Unicode.\n
	/// If set to \c VARIANT_TRUE, the conversion is performed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DoOEMConversion, get_CharacterConversion, get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647473.aspx">CharToOem</a>
	virtual HRESULT STDMETHODCALLTYPE get_DoOEMConversion(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DoOEMConversion property</em>
	///
	/// Sets whether the contained edit control's text is converted from the Windows character set to
	/// the OEM character set and then back to the Windows character set. Such a conversion ensures proper
	/// character conversion when the application calls the \c CharToOem function to convert a Windows string
	/// in the contained edit control to OEM characters. This property is most useful if the contained edit
	/// control contains file names that will be used on file systems that do not support Unicode.\n
	/// If set to \c VARIANT_TRUE, the conversion is performed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DoOEMConversion, put_CharacterConversion, put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647473.aspx">CharToOem</a>
	virtual HRESULT STDMETHODCALLTYPE put_DoOEMConversion(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, get_ReadOnly
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, put_ReadOnly
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FirstVisibleChar property</em>
	///
	/// Retrieves the zero-based index of the first visible character in the contained edit control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_FirstVisibleChar(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the control's content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont, get_ForeColor, get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont, put_ForeColor, put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Font, put_Font, put_UseSystemFont, put_ForeColor, put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ForeColor, get_BackColor, get_DisabledForeColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ForeColor, put_BackColor, put_DisabledForeColor
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleHeight property</em>
	///
	/// Retrieves the height (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the contained edit control's text is governed by the dimensions of its window
	/// rectangle and its formatting rectangle. The formatting rectangle is a construct maintained by the
	/// system for formatting the text displayed in the window rectangle. When the contained edit control is
	/// first displayed, the two rectangles are identical on the screen. An application can make the
	/// formatting rectangle larger than the window rectangle (thereby limiting the visibility of the
	/// contained edit control's text) or smaller than the window rectangle (thereby creating extra white
	/// space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FormattingRectangleLeft, get_FormattingRectangleTop, get_FormattingRectangleWidth
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleLeft property</em>
	///
	/// Retrieves the distance (in pixels) between the left borders of the control's formatting rectangle and
	/// its client area.\n
	/// The visibility of the contained edit control's text is governed by the dimensions of its window
	/// rectangle and its formatting rectangle. The formatting rectangle is a construct maintained by the
	/// system for formatting the text displayed in the window rectangle. When the contained edit control is
	/// first displayed, the two rectangles are identical on the screen. An application can make the
	/// formatting rectangle larger than the window rectangle (thereby limiting the visibility of the
	/// contained edit control's text) or smaller than the window rectangle (thereby creating extra white
	/// space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FormattingRectangleHeight, get_FormattingRectangleTop, get_FormattingRectangleWidth
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleLeft(OLE_XPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleTop property</em>
	///
	/// Retrieves the distance (in pixels) between the upper borders of the control's formatting rectangle
	/// and its client area.\n
	/// The visibility of the contained edit control's text is governed by the dimensions of its window
	/// rectangle and its formatting rectangle. The formatting rectangle is a construct maintained by the
	/// system for formatting the text displayed in the window rectangle. When the contained edit control is
	/// first displayed, the two rectangles are identical on the screen. An application can make the
	/// formatting rectangle larger than the window rectangle (thereby limiting the visibility of the
	/// contained edit control's text) or smaller than the window rectangle (thereby creating extra white
	/// space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FormattingRectangleHeight, get_FormattingRectangleLeft, get_FormattingRectangleWidth
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleTop(OLE_YPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleWidth property</em>
	///
	/// Retrieves the width (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the contained edit control's text is governed by the dimensions of its window
	/// rectangle and its formatting rectangle. The formatting rectangle is a construct maintained by the
	/// system for formatting the text displayed in the window rectangle. When the contained edit control is
	/// first displayed, the two rectangles are identical on the screen. An application can make the
	/// formatting rectangle larger than the window rectangle (thereby limiting the visibility of the
	/// contained edit control's text) or smaller than the window rectangle (thereby creating extra white
	/// space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FormattingRectangleHeight, get_FormattingRectangleLeft, get_FormattingRectangleTop
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c GroupDigits property</em>
	///
	/// Retrieves whether digits are grouped according to the locale settings when displaying the current
	/// value in the contained edit control. If set to \c VARIANT_TRUE, digits are grouped; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property has no effect if the \c AutomaticallySetText property is set to
	///          \c VARIANT_FALSE.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa put_GroupDigits, get_CurrentValue, get_AutomaticallySetText
	virtual HRESULT STDMETHODCALLTYPE get_GroupDigits(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c GroupDigits property</em>
	///
	/// Sets whether digits are grouped according to the locale settings when displaying the current
	/// value in the contained edit control. If set to \c VARIANT_TRUE, digits are grouped; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property has no effect if the \c AutomaticallySetText property is set to
	///          \c VARIANT_FALSE.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa get_GroupDigits, put_CurrentValue, put_AutomaticallySetText
	virtual HRESULT STDMETHODCALLTYPE put_GroupDigits(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c HAlignment property</em>
	///
	/// Retrieves the horizontal alignment of the control's content. Any of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention On Windows XP, changing this property destroys and recreates the edit control window.
	///
	/// \if UNICODE
	///   \sa put_HAlignment, get_UpDownPosition, get_Text, EditCtlsLibU::HAlignmentConstants
	/// \else
	///   \sa put_HAlignment, get_UpDownPosition, get_Text, EditCtlsLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HAlignment(HAlignmentConstants* pValue);
	/// \brief <em>Sets the \c HAlignment property</em>
	///
	/// Sets the horizontal alignment of the control's content. Any of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention On Windows XP, changing this property destroys and recreates the edit control window.
	///
	/// \if UNICODE
	///   \sa get_HAlignment, put_UpDownPosition, put_Text, EditCtlsLibU::HAlignmentConstants
	/// \else
	///   \sa get_HAlignment, put_UpDownPosition, put_Text, EditCtlsLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c HotTracking property</em>
	///
	/// Retrieves whether the arrows get highlighted if the mouse cursor is moved over them. If set to
	/// \c VARIANT_TRUE, the arrows get highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored for themed up down controls.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa put_HotTracking
	virtual HRESULT STDMETHODCALLTYPE get_HotTracking(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c HotTracking property</em>
	///
	/// Sets whether the arrows get highlighted if the mouse cursor is moved over them. If set to
	/// \c VARIANT_TRUE, the arrows get highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored for themed up down controls.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa get_HotTracking
	virtual HRESULT STDMETHODCALLTYPE put_HotTracking(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWndEdit, get_hWndUpDown, Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndEdit property</em>
	///
	/// Retrieves the window handle of the control's contained edit window.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWnd, get_hWndUpDown, Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWndEdit(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndUpDown property</em>
	///
		/// Retrieves the window handle of the control's contained up down window.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWnd, get_hWndEdit, Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWndUpDown(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c IMEMode property</em>
	///
	/// Retrieves the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IMEMode, EditCtlsLibU::IMEModeConstants
	/// \else
	///   \sa put_IMEMode, EditCtlsLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IMEMode(IMEModeConstants* pValue);
	/// \brief <em>Sets the \c IMEMode property</em>
	///
	/// Sets the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IMEMode, EditCtlsLibU::IMEModeConstants
	/// \else
	///   \sa get_IMEMode, EditCtlsLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IMEMode(IMEModeConstants newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c LeftMargin property</em>
	///
	/// Retrieves the width (in pixels) of the contained edit control's left margin. If set to -1, a value,
	/// that depends on the contained edit control's font, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LeftMargin, get_RightMargin, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_LeftMargin(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c LeftMargin property</em>
	///
	/// Sets the width (in pixels) of the contained edit control's left margin. If set to -1, a value,
	/// that depends on the contained edit control's font, is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LeftMargin, put_RightMargin, put_Font
	virtual HRESULT STDMETHODCALLTYPE put_LeftMargin(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c Maximum property</em>
	///
	/// Retrieves the maximum value that the user can select.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Maximum, get_Minimum, get_CurrentValue, get_WrapAtBoundaries
	virtual HRESULT STDMETHODCALLTYPE get_Maximum(LONG* pValue);
	/// \brief <em>Sets the \c Maximum property</em>
	///
	/// Sets the maximum value that the user can select.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Maximum, put_Minimum, put_CurrentValue, put_WrapAtBoundaries
	virtual HRESULT STDMETHODCALLTYPE put_Maximum(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MaxTextLength property</em>
	///
	/// Retrieves the maximum number of characters, that the user can type into the contained edit control.
	/// If set to -1, the system's default setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c CurrentValue property may exceed this limit.
	///
	/// \sa put_MaxTextLength, get_Text, get_CurrentValue, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE get_MaxTextLength(LONG* pValue);
	/// \brief <em>Sets the \c MaxTextLength property</em>
	///
	/// Sets the maximum number of characters, that the user can type into the contained edit control.
	/// If set to -1, the system's default setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c CurrentValue property may exceed this limit.
	///
	/// \sa get_MaxTextLength, put_Text, put_CurrentValue, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE put_MaxTextLength(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Minimum property</em>
	///
	/// Retrieves the minimum value that the user can select.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Minimum, get_Maximum, get_CurrentValue, get_WrapAtBoundaries
	virtual HRESULT STDMETHODCALLTYPE get_Minimum(LONG* pValue);
	/// \brief <em>Sets the \c Minimum property</em>
	///
	/// Sets the minimum value that the user can select.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Minimum, put_Maximum, put_CurrentValue, put_WrapAtBoundaries
	virtual HRESULT STDMETHODCALLTYPE put_Minimum(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Modified property</em>
	///
	/// Retrieves a flag indicating whether the contained edit control's content has changed. A value of
	/// \c VARIANT_TRUE stands for changed content, a value of \c VARIANT_FALSE for unchanged content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Modified, get_Text, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE get_Modified(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Modified property</em>
	///
	/// Sets a flag indicating whether the contained edit control's content has changed. A value of
	/// \c VARIANT_TRUE stands for changed content, a value of \c VARIANT_FALSE for unchanged content.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Modified, put_Text, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE put_Modified(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, EditCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, EditCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, EditCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MousePointer, get_MouseIcon, EditCtlsLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, EditCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MousePointer, put_MouseIcon, EditCtlsLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, EditCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Orientation property</em>
	///
	/// Retrieves the control's orientation. Any of the values defined by the \c OrientationConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \if UNICODE
	///   \sa put_Orientation, get_UpDownPosition, EditCtlsLibU::OrientationConstants
	/// \else
	///   \sa put_Orientation, get_UpDownPosition, EditCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Orientation(OrientationConstants* pValue);
	/// \brief <em>Sets the \c Orientation property</em>
	///
	/// Sets the control's orientation. Any of the values defined by the \c OrientationConstants enumeration
	/// is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \if UNICODE
	///   \sa get_Orientation, put_UpDownPosition, EditCtlsLibU::OrientationConstants
	/// \else
	///   \sa get_Orientation, put_UpDownPosition, EditCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Orientation(OrientationConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessArrowKeys property</em>
	///
	/// Retrieves whether the current value is incremented or decremented if the user presses the [UP] or
	/// [DOWN] arrow key. If set to \c VARIANT_TRUE, the current value is changed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa put_ProcessArrowKeys, get_CurrentValue, get_ProcessContextMenuKeys
	virtual HRESULT STDMETHODCALLTYPE get_ProcessArrowKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessArrowKeys property</em>
	///
	/// Sets whether the current value is incremented or decremented if the user presses the [UP] or
	/// [DOWN] arrow key. If set to \c VARIANT_TRUE, the current value is changed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa get_ProcessArrowKeys, put_CurrentValue, put_ProcessContextMenuKeys
	virtual HRESULT STDMETHODCALLTYPE put_ProcessArrowKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, get_ProcessArrowKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, put_ProcessArrowKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c ReadOnlyTextBox property</em>
	///
	/// Retrieves whether the contained edit control accepts user input, that would change its content. If
	/// set to \c VARIANT_FALSE, such user input is accepted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE, the user is still able to change the content by
	///          clicking the up down control's buttons or using the arrow keys on the keyboard.
	///
	/// \sa put_ReadOnlyTextBox, get_Enabled, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ReadOnlyTextBox(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ReadOnlyTextBox property</em>
	///
	/// Sets whether the contained edit control accepts user input, that would change its content. If
	/// set to \c VARIANT_FALSE, such user input is accepted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE, the user is still able to change the content by
	///          clicking the up down control's buttons or using the arrow keys on the keyboard.
	///
	/// \sa get_ReadOnlyTextBox, put_Enabled, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_ReadOnlyTextBox(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RegisterForOLEDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RegisterForOLEDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RightMargin property</em>
	///
	/// Retrieves the width (in pixels) of the contained edit control's right margin. If set to -1, a value,
	/// that depends on the contained edit control's font, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RightMargin, get_LeftMargin, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_RightMargin(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c RightMargin property</em>
	///
	/// Sets the width (in pixels) of the contained edit control's right margin. If set to -1, a value,
	/// that depends on the contained edit control's font, is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RightMargin, put_LeftMargin, put_Font
	virtual HRESULT STDMETHODCALLTYPE put_RightMargin(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, get_IMEMode, Raise_WritingDirectionChanged, EditCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, get_IMEMode, Raise_WritingDirectionChanged, EditCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Setting or clearing the \c rtlLayout flag destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, put_IMEMode, Raise_WritingDirectionChanged, EditCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, put_IMEMode, Raise_WritingDirectionChanged, EditCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c SelectedText property</em>
	///
	/// Retrieves the currently selected text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa GetSelection, SetSelection, ReplaceSelectedText, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_SelectedText(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the contained edit control's content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, get_CurrentValue, get_AutomaticallySetText, get_DoOEMConversion, get_MaxTextLength,
	///     get_AcceptNumbersOnly, get_CueBanner, get_HAlignment, get_ForeColor, get_Font, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the contained edit control's content.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, put_CurrentValue, put_AutomaticallySetText, put_DoOEMConversion, put_MaxTextLength,
	///     put_AcceptNumbersOnly, put_CueBanner, put_HAlignment, put_ForeColor, put_Font, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c UpDownPosition property</em>
	///
	/// Retrieves the location of the contained up down control relative to the contained edit control. Any
	/// of the values defined by the \c UpDownPositionConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \if UNICODE
	///   \sa put_UpDownPosition, get_HAlignment, get_Orientation, EditCtlsLibU::UpDownPositionConstants
	/// \else
	///   \sa put_UpDownPosition, get_HAlignment, get_Orientation, EditCtlsLibA::UpDownPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_UpDownPosition(UpDownPositionConstants* pValue);
	/// \brief <em>Sets the \c UpDownPosition property</em>
	///
	/// Sets the location of the contained up down control relative to the contained edit control. Any
	/// of the values defined by the \c UpDownPositionConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \if UNICODE
	///   \sa get_UpDownPosition, put_HAlignment, put_Orientation, EditCtlsLibU::UpDownPositionConstants
	/// \else
	///   \sa get_UpDownPosition, put_HAlignment, put_Orientation, EditCtlsLibA::UpDownPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_UpDownPosition(UpDownPositionConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseSystemFont, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c WordBreakFunction property</em>
	///
	/// Retrieves the function that is responsible to tell the contained edit control where a word starts
	/// and where it ends. This property takes the address of a function having the following signature:\n
	/// \code
	///   int CALLBACK FindWorkBreak(LPTSTR pText, int startPosition, int textLength, int flags);
	/// \endcode
	/// The \c pText argument is a pointer to the contained edit control's text.\n
	/// The \c startPosition argument specifies the (zero-based) position within the text, at which the
	/// function should begin checking for a word break.\n
	/// The \c textLength argument specifies the length of the text pointed to by \c pText in characters.\n
	/// The \c flags argument specifies the action to be taken by the function. This can be one of the
	/// following values:
	/// - \c WB_ISDELIMITER Check whether the character at the specified position is a delimiter.
	/// - \c WB_LEFT Find the beginning of a word to the left of the specified position.
	/// - \c WB_RIGHT Find the beginning of a word to the right of the specified position. This is useful in
	///   right-aligned edit controls.
	///
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is a delimiter, the function must return \c TRUE.\n
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is not a delimiter, the function must return \c FALSE.\n
	/// If the \c flags parameter specifies \c WB_LEFT or \c WB_RIGHT, the function must return the
	/// (zero-based) index to the beginning of a word in the specified text.\n\n
	/// If this property is set to 0, the system's internal function is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_WordBreakFunction, get_Text, get_HAlignment,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672125.aspx">EditWordBreakProc</a>
	virtual HRESULT STDMETHODCALLTYPE get_WordBreakFunction(LONG* pValue);
	/// \brief <em>Sets the \c WordBreakFunction property</em>
	///
	/// Sets the function that is responsible to tell the contained edit control where a word starts
	/// and where it ends. This property takes the address of a function having the following signature:\n
	/// \code
	///   int CALLBACK FindWorkBreak(LPTSTR pText, int startPosition, int textLength, int flags);
	/// \endcode
	/// The \c pText argument is a pointer to the contained edit control's text.\n
	/// The \c startPosition argument specifies the (zero-based) position within the text, at which the
	/// function should begin checking for a word break.\n
	/// The \c textLength argument specifies the length of the text pointed to by \c pText in characters.\n
	/// The \c flags argument specifies the action to be taken by the function. This can be one of the
	/// following values:
	/// - \c WB_ISDELIMITER Check whether the character at the specified position is a delimiter.
	/// - \c WB_LEFT Find the beginning of a word to the left of the specified position.
	/// - \c WB_RIGHT Find the beginning of a word to the right of the specified position. This is useful in
	///   right-aligned edit controls.
	///
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is a delimiter, the function must return \c TRUE.\n
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is not a delimiter, the function must return \c FALSE.\n
	/// If the \c flags parameter specifies \c WB_LEFT or \c WB_RIGHT, the function must return the
	/// (zero-based) index to the beginning of a word in the specified text.\n\n
	/// If this property is set to 0, the system's internal function is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_WordBreakFunction, put_Text, put_HAlignment,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672125.aspx">EditWordBreakProc</a>
	virtual HRESULT STDMETHODCALLTYPE put_WordBreakFunction(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c WrapAtBoundaries property</em>
	///
	/// Retrieves whether the current value "wraps" to the opposite end of the range if it is incremented
	/// or decremented beyond the ending or beginning of the range. If set to \c VARIANT_TRUE, the value
	/// "wraps"; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa put_WrapAtBoundaries, get_Maximum, get_Minimum, get_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE get_WrapAtBoundaries(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c WrapAtBoundaries property</em>
	///
	/// Sets whether the current value "wraps" to the opposite end of the range if it is incremented
	/// or decremented beyond the ending or beginning of the range. If set to \c VARIANT_TRUE, the value
	/// "wraps"; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the contained up down window.
	///
	/// \sa get_WrapAtBoundaries, put_Maximum, put_Minimum, put_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE put_WrapAtBoundaries(VARIANT_BOOL newValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Determines whether there are any actions in the contained edit control's undo queue</em>
	///
	/// \param[out] pValue \c VARIANT_TRUE if there are actions in the undo queue; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Undo, EmptyUndoBuffer
	virtual HRESULT STDMETHODCALLTYPE CanUndo(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the specified character's position in client coordinates</em>
	///
	/// \param[in] characterIndex The zero-based index of the character within the contained edit control,
	///            for which to retrieve the position. If the specified index is greater than the index of
	///            the last character in the contained edit control, the function fails.
	/// \param[out] pX The x-coordinate (in pixels) of the character relative to the contained edit control's
	///             upper-left corner.
	/// \param[out] pY The y-coordinate (in pixels) of the character relative to the contained edit control's
	///             upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PositionToCharIndex
	virtual HRESULT STDMETHODCALLTYPE CharIndexToPosition(LONG characterIndex, OLE_XPOS_PIXELS* pX = NULL, OLE_YPOS_PIXELS* pY = NULL);
	/// \brief <em>Clears the contained edit control's undo queue</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanUndo, Undo
	virtual HRESULT STDMETHODCALLTYPE EmptyUndoBuffer(void);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Retrieves the current selection's start and end</em>
	///
	/// Retrieves the zero-based character indices of the current selection's start and end.
	///
	/// \param[out] pSelectionStart The zero-based index of the character at which the selection starts.
	/// \param[out] pSelectionEnd The zero-based index of the first unselected character after the end of the
	///             selection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetSelection, ReplaceSelectedText, get_SelectedText
	virtual HRESULT STDMETHODCALLTYPE GetSelection(LONG* pSelectionStart = NULL, LONG* pSelectionEnd = NULL);
	/// \brief <em>Hides any balloon tips associated with the contained edit control</em>
	///
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa ShowBalloonTip
	virtual HRESULT STDMETHODCALLTYPE HideBalloonTip(VARIANT_BOOL* pSucceeded);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pHitTestDetails Receives a value specifying the exact part of the control the specified
	///             point lies in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa EditCtlsLibA::HitTestConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails);
	/// \brief <em>Determines whether the text displayed by the contained edit control is entirely visible or truncated</em>
	///
	/// \param[out] pValue \c VARIANT_TRUE if not all of the text is visible; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE IsTextTruncated(VARIANT_BOOL* pValue);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Retrieves the character closest to the specified position</em>
	///
	/// Retrieves the zero-based index of the character nearest the specified position.
	///
	/// \param[in] x The x-coordinate (in pixels) of the position to retrieve the nearest character for. It
	///            is relative to the contained edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the position to retrieve the nearest character for. It
	///            is relative to the contained edit control's upper-left corner.
	/// \param[out] pCharacterIndex The zero-based index of the character within the contained edit control,
	///             that is nearest to the specified position. If the specified point is beyond the last
	///             character in the contained edit control, this value indicates the last character in the
	///             contained edit control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If a point outside the bounds of the contained edit control is passed, the function fails.
	///
	/// \sa CharIndexToPosition
	virtual HRESULT STDMETHODCALLTYPE PositionToCharIndex(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG* pCharacterIndex = NULL);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Replaces the currently selected text</em>
	///
	/// Replaces the contained edit control's currently selected text.
	///
	/// \param[in] replacementText The text that replaces the currently selected text.
	/// \param[in] undoable If \c VARIANT_TRUE, this action is inserted into the contained edit control's
	///            undo queue; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetSelection, SetSelection, get_SelectedText
	virtual HRESULT STDMETHODCALLTYPE ReplaceSelectedText(BSTR replacementText, VARIANT_BOOL undoable = VARIANT_FALSE);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Scrolls the contained edit control so that the caret is visible</em>
	///
	/// Ensures that the contained edit control's caret is visible by scrolling the contained edit control
	/// if necessary.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutoScrolling
	virtual HRESULT STDMETHODCALLTYPE ScrollCaretIntoView(void);
	/// \brief <em>Sets the selection's start and end</em>
	///
	/// Sets the zero-based character indices of the selection's start and end.
	///
	/// \param[in] selectionStart The zero-based index of the character at which the selection starts. If set
	///            to -1, the current selection is cleared.
	/// \param[in] selectionEnd The zero-based index of the first unselected character after the end of the
	///            selection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks To select all text in the contained edit control, set \c selectionStart to 0 and
	///          \c selectionEnd to -1.
	///
	/// \sa GetSelection, ReplaceSelectedText, put_SelectedText
	virtual HRESULT STDMETHODCALLTYPE SetSelection(LONG selectionStart, LONG selectionEnd);
	/// \brief <em>Displays a balloon tip associated with the contained edit control</em>
	///
	/// \param[in] title The title of the balloon tip to display.
	/// \param[in] text The balloon tip text to display.
	/// \param[in] icon The icon of the balloon tip to display. Any of the values defined by the
	///            \c BalloonTipIconConstants enumeration is valid.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \if UNICODE
	///   \sa HideBalloonTip, EditCtlsLibU::BalloonTipIconConstants
	/// \else
	///   \sa HideBalloonTip, EditCtlsLibA::BalloonTipIconConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE ShowBalloonTip(BSTR title, BSTR text, BalloonTipIconConstants icon = btiNone, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Undoes the last action in the contained edit control's undo queue</em>
	///
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanUndo, EmptyUndoBuffer
	virtual HRESULT STDMETHODCALLTYPE Undo(VARIANT_BOOL* pSucceeded);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CComCompositeControl::Create() and is
	/// used to raise the \c RecreatedControlWindow event.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] createParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnInitDialog, Raise_RecreatedControlWindow
	HWND Create(HWND hWndParent, RECT& rect, LPARAM createParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/056hw3hs.aspx">CComControlBase::OnDraw</a>
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnInitDialog, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	// \param[in] languageID The locale identifier identifying the language in which name should be
	//            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties, StringProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CTLCOLOREDIT handler</em>
	///
	/// Will be called if the contained edit control asks its parent window to configure the specified device
	/// context for drawing the control, i. e. to setup the colors and brushes.
	/// We use this handler for the \c BackColor and \c ForeColor properties.
	///
	/// \sa get_BackColor, get_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672119.aspx">WM_CTLCOLOREDIT</a>
	LRESULT OnCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CTLCOLORSTATIC handler</em>
	///
	/// Will be called if the contained edit control asks its parent window to configure the specified device
	/// context for drawing the control, i. e. to setup the colors and brushes.
	/// We use this handler for proper theming support and for the \c Enabled and \c ReadOnlyTextBox
	/// properties.
	///
	/// \sa get_DisabledBackColor, get_DisabledForeColor, get_Enabled, get_ReadOnlyTextBox,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651165.aspx">WM_CTLCOLORSTATIC</a>
	LRESULT OnCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnInitDialog, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_INITDIALOG handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645428.aspx">WM_INITDIALOG</a>
	LRESULT OnInitDialog(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnEditSetFocus, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_*SCROLL handler</em>
	///
	/// Will be called if the contained up down control's value has changed.
	/// We use this handler to raise the \c ValueChanged event.
	///
	/// \sa OnDeltaPosNotification, Raise_ValueChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651283.aspx">WM_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651284.aspx">WM_VSCROLL</a>
	LRESULT OnScroll(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa OnEditSetRedraw, OnUpDownSetRedraw, get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnSetRedraw(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETTINGCHANGE handler</em>
	///
	/// Will be called if a system setting was changed.
	/// We use this handler to update our appearance.
	///
	/// \attention This message is posted to top-level windows only, so actually we'll never receive it.
	///
	/// \sa OnThemeChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms725497.aspx">WM_SETTINGCHANGE</a>
	LRESULT OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to update our appearance.
	///
	/// \sa OnSettingChange, Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler for the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnEditKeyDown, OnEditIMEChar, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnEditChar(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the contained edit control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnUpDownContextMenu, OnEditRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnEditContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_IME_CHAR handler</em>
	///
	/// Sent by the input method editor (IME) when a character has been composited while the contained edit
	/// control has the keyboard focus.
	/// We use this handler to make more IME implementations (namely the one for emojis) work with ANSI
	/// applications like all VB6 applications.
	///
	/// \sa OnEditChar,
	///     <a href="https://docs.microsoft.com/en-us/windows/win32/intl/wm-ime-char">WM_IME_CHAR</a>
	LRESULT OnEditIMEChar(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_INPUTLANGCHANGE handler</em>
	///
	/// Will be called after an application's input language has been changed.
	/// We use this handler to update the IME mode of the control.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632629.aspx">WM_INPUTLANGCHANGE</a>
	LRESULT OnEditInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the contained edit control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnEditKeyUp, OnEditChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnEditKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the contained edit control has the keyboard
	/// focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnEditKeyDown, OnEditChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnEditKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained edit control's client area using the
	/// left mouse button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnUpDownLButtonDblClk, OnEditMButtonDblClk, OnEditRButtonDblClk, OnEditXButtonDblClk,
	///     Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnEditLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the contained edit control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnUpDownLButtonDown, OnEditMButtonDown, OnEditRButtonDown, OnEditXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnEditLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the contained edit control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnUpDownLButtonUp, OnEditMButtonUp, OnEditRButtonUp, OnEditXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnEditLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained edit control's client area using the
	/// middle mouse button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnUpDownMButtonDblClk, OnEditLButtonDblClk, OnEditRButtonDblClk, OnEditXButtonDblClk,
	///     Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnEditMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the contained edit control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnUpDownMButtonDown, OnEditLButtonDown, OnEditRButtonDown, OnEditXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnEditMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the contained edit control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnUpDownMButtonUp, OnEditLButtonUp, OnEditRButtonUp, OnEditXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnEditMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the contained edit control's client area for
	/// a previously specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa OnUpDownMouseHover, Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnEditMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the contained edit control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa OnUpDownMouseLeave, Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnEditMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the contained edit
	/// control's client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnUpDownMouseMove, OnEditLButtonDown, OnEditLButtonUp, OnEditMouseWheel, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnEditMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// contained edit control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnUpDownMouseWheel, OnEditMouseMove, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnEditMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained edit control's client area using the
	/// right mouse button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnUpDownRButtonDblClk, OnEditLButtonDblClk, OnEditMButtonDblClk, OnEditXButtonDblClk,
	///     Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnEditRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the contained edit control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnUpDownRButtonDown, OnEditLButtonDown, OnEditMButtonDown, OnEditXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnEditRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the contained edit control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnUpDownRButtonUp, OnEditLButtonUp, OnEditMButtonUp, OnEditXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnEditRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called if the contained edit control receives the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnSetFocus, OnUpDownSetFocus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnEditSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the contained edit control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnEditSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the contained edit control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa OnSetRedraw, OnUpDownSetRedraw, get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnEditSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTEXT handler</em>
	///
	/// Will be called if the contained edit control's text shall be changed.
	/// We use this handler for better data-binding support.
	///
	/// \sa get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632644.aspx">WM_SETTEXT</a>
	LRESULT OnEditSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained edit control's client area using one of
	/// the extended mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnUpDownXButtonDblClk, OnEditLButtonDblClk, OnEditMButtonDblClk, OnEditRButtonDblClk,
	///     Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnEditXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the contained edit control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnUpDownXButtonDown, OnEditLButtonDown, OnEditMButtonDown, OnEditRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnEditXButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the contained edit control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnUpDownXButtonUp, OnEditLButtonUp, OnEditMButtonUp, OnEditRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnEditXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c EM_SETCUEBANNER handler</em>
	///
	/// Will be called if the contained edit control's textual cue shall be set.
	/// We use this handler to synchronize our settings.
	///
	/// \sa get_DisplayCueBannerOnFocus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672093.aspx">EM_SETCUEBANNER</a>
	LRESULT OnSetCueBanner(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the contained up down control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnEditContextMenu, OnUpDownRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnUpDownContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained up down control's client area using the
	/// left mouse button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnEditLButtonDblClk, OnUpDownMButtonDblClk, OnUpDownRButtonDblClk, OnUpDownXButtonDblClk,
	///     Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnUpDownLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the contained up down control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnEditLButtonDown, OnUpDownMButtonDown, OnUpDownRButtonDown, OnUpDownXButtonDown,
	///     Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnUpDownLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the contained up down control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnEditLButtonUp, OnUpDownMButtonUp, OnUpDownRButtonUp, OnUpDownXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnUpDownLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained up down control's client area using the
	/// middle mouse button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnEditMButtonDblClk, OnUpDownLButtonDblClk, OnUpDownRButtonDblClk, OnUpDownXButtonDblClk,
	///     Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnUpDownMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the contained up down control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnEditMButtonDown, OnUpDownLButtonDown, OnUpDownRButtonDown, OnUpDownXButtonDown,
	///     Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnUpDownMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the contained up down control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnEditMButtonUp, OnUpDownLButtonUp, OnUpDownRButtonUp, OnUpDownXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnUpDownMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the contained up down control's client area
	/// for a previously specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa OnEditMouseHover, Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnUpDownMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the contained up down control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa OnEditMouseLeave, Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnUpDownMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the contained up
	/// down control's client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnEditMouseMove, OnUpDownLButtonDown, OnUpDownLButtonUp, OnUpDownMouseWheel, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnUpDownMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// contained up down control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnEditMouseWheel, OnUpDownMouseMove, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnUpDownMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained up down control's client area using the
	/// right mouse button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnEditRButtonDblClk, OnUpDownLButtonDblClk, OnUpDownMButtonDblClk, OnUpDownXButtonDblClk,
	///     Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnUpDownRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the contained up down control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnEditRButtonDown, OnUpDownLButtonDown, OnUpDownMButtonDown, OnUpDownXButtonDown,
	///     Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnUpDownRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the contained up down control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnEditRButtonUp, OnUpDownLButtonUp, OnUpDownMButtonUp, OnUpDownXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnUpDownRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called if the contained up down control receives the keyboard focus.
	/// We use this handler to make the control selectable using [SHIFT]+[TAB].
	///
	/// \sa OnSetFocus, OnEditSetFocus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnUpDownSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the contained up down control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa OnSetRedraw, OnEditSetRedraw, get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnUpDownSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the contained up down control's client area using one
	/// of the extended mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnEditXButtonDblClk, OnUpDownLButtonDblClk, OnUpDownMButtonDblClk, OnUpDownRButtonDblClk,
	///     Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnUpDownXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the contained up down control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnEditXButtonDown, OnUpDownLButtonDown, OnUpDownMButtonDown, OnUpDownRButtonDown,
	///     Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnUpDownXButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the contained up down control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnEditXButtonUp, OnUpDownLButtonUp, OnUpDownMButtonUp, OnUpDownRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnUpDownXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c UDM_SETACCEL handler</em>
	///
	/// Will be called if the contained up down control's accelerators shall be redefined.
	/// We use this handler to raise the \c ChangingAccelerators and \c ChangedAccelerators events.
	///
	/// \sa Raise_ChangingAccelerators, Raise_ChangedAccelerators,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649988.aspx">UDM_SETACCEL</a>
	LRESULT OnSetAccel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c UDM_SETRANGE* handler</em>
	///
	/// Will be called if the contained up down control's value range shall be redefined.
	/// We use this handler for the \c AutomaticallyCorrectValue property.
	///
	/// \sa get_AutomaticallyCorrectValue,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650013.aspx">UDM_SETRANGE</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650014.aspx">UDM_SETRANGE32</a>
	LRESULT OnSetRange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c UDN_DELTAPOS handler</em>
	///
	/// Will be called if the control is notified, that the contained up down control's value is about to
	/// change.
	/// We use this handler to raise the \c ValueChanging event.
	///
	/// \sa OnChange, Raise_ValueChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649979.aspx">UDN_DELTAPOS</a>
	LRESULT OnDeltaPosNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c EN_ALIGN_*_EC handler</em>
	///
	/// Will be called if the control is notified, that the user has changed the contained edit control's
	/// writing direction.
	/// We use this handler to raise the \c WritingDirectionChanged event.
	///
	/// \sa get_RightToLeft, Raise_WritingDirectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672108.aspx">EN_ALIGN_LTR_EC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672109.aspx">EN_ALIGN_RTL_EC</a>
	LRESULT OnAlign(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_CHANGE handler</em>
	///
	/// Will be called if the control is notified, that the contained edit control's text has changed.
	/// We use this handler to raise the \c TextChanged event.
	///
	/// \sa OnDeltaPosNotification, Raise_TextChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672110.aspx">EN_CHANGE</a>
	LRESULT OnChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_ERRSPACE handler</em>
	///
	/// Will be called if the control is notified, that the contained edit control couldn't allocate enough
	/// memory to meet a specific request.
	/// We use this handler to raise the \c OutOfMemory event.
	///
	/// \sa Raise_OutOfMemory,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672112.aspx">EN_ERRSPACE</a>
	LRESULT OnErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_KILLFOCUS handler</em>
	///
	/// Will be called if the control is notified, that the contained edit control has lost the keyboard
	/// focus.
	/// We use this handler for the \c AutomaticallyCorrectValue property.
	///
	/// \sa get_AutomaticallyCorrectValue,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672114.aspx">EN_KILLFOCUS</a>
	LRESULT OnKillFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_MAXTEXT handler</em>
	///
	/// Will be called if the control is notified, that the text, that was entered into the contained edit
	/// control, got truncated.
	/// We use this handler to raise the \c TruncatedText event.
	///
	/// \sa Raise_TruncatedText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672115.aspx">EN_MAXTEXT</a>
	LRESULT OnMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_SETFOCUS handler</em>
	///
	/// Will be called if the control is notified, that the contained edit control has gained the keyboard
	/// focus.
	/// We use this handler to initialize IME.
	///
	/// \sa OnEditSetFocus, OnUpDownSetFocus, get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672116.aspx">EN_SETFOCUS</a>
	LRESULT OnSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND hWnd, BOOL& wasHandled);
	/// \brief <em>\c EN_UPDATE handler</em>
	///
	/// Will be called if the control is notified, that the contained edit control's content is about to be
	/// drawn.
	/// We use this handler to raise the \c BeforeDrawText event.
	///
	/// \sa Raise_BeforeDrawText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672117.aspx">EN_UPDATE</a>
	LRESULT OnUpdate(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c BeforeDrawText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_BeforeDrawText,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::BeforeDrawText
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_BeforeDrawText,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::BeforeDrawText
	/// \endif
	inline HRESULT Raise_BeforeDrawText(void);
	/// \brief <em>Raises the \c ChangedAccelerators event</em>
	///
	/// \param[in] pAccelerators A collection of the new accelerators.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ChangedAccelerators,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::ChangedAccelerators, Raise_ChangingAccelerators
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ChangedAccelerators,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::ChangedAccelerators, Raise_ChangingAccelerators
	/// \endif
	inline HRESULT Raise_ChangedAccelerators(IUpDownAccelerators* pAccelerators);
	/// \brief <em>Raises the \c ChangingAccelerators event</em>
	///
	/// \param[in] pAccelerators A collection of the new accelerators.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort redefining; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ChangingAccelerators,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::ChangingAccelerators, Raise_ChangedAccelerators
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ChangingAccelerators,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::ChangingAccelerators, Raise_ChangedAccelerators
	/// \endif
	inline HRESULT Raise_ChangingAccelerators(IVirtualUpDownAccelerators* pAccelerators, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_Click, EditCtlsLibU::_IUpDownTextBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_Click, EditCtlsLibA::_IUpDownTextBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the control from
	///                showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ContextMenu, EditCtlsLibU::_IUpDownTextBoxEvents::ContextMenu,
	///       Raise_RClick, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ContextMenu, EditCtlsLibA::_IUpDownTextBoxEvents::ContextMenu,
	///       Raise_RClick, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_DblClick, EditCtlsLibU::_IUpDownTextBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_DblClick, EditCtlsLibA::_IUpDownTextBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_DestroyedControlWindow,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_DestroyedControlWindow,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_KeyDown, EditCtlsLibU::_IUpDownTextBoxEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_KeyDown, EditCtlsLibA::_IUpDownTextBoxEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_KeyPress, EditCtlsLibU::_IUpDownTextBoxEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_KeyPress, EditCtlsLibA::_IUpDownTextBoxEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \endif
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_KeyUp, EditCtlsLibU::_IUpDownTextBoxEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_KeyUp, EditCtlsLibA::_IUpDownTextBoxEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MClick, EditCtlsLibU::_IUpDownTextBoxEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MClick, EditCtlsLibA::_IUpDownTextBoxEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MDblClick, EditCtlsLibU::_IUpDownTextBoxEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MDblClick, EditCtlsLibA::_IUpDownTextBoxEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseDown, EditCtlsLibU::_IUpDownTextBoxEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseDown, EditCtlsLibA::_IUpDownTextBoxEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseEnter, EditCtlsLibU::_IUpDownTextBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseEnter, EditCtlsLibA::_IUpDownTextBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseHover, EditCtlsLibU::_IUpDownTextBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseHover, EditCtlsLibA::_IUpDownTextBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseLeave, EditCtlsLibU::_IUpDownTextBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseLeave, EditCtlsLibA::_IUpDownTextBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseMove, EditCtlsLibU::_IUpDownTextBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseMove, EditCtlsLibA::_IUpDownTextBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseUp, EditCtlsLibU::_IUpDownTextBoxEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseUp, EditCtlsLibA::_IUpDownTextBoxEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseWheel, EditCtlsLibU::_IUpDownTextBoxEvents::MouseWheel,
	///       Raise_MouseMove, EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_MouseWheel, EditCtlsLibA::_IUpDownTextBoxEvents::MouseWheel,
	///       Raise_MouseMove, EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragDrop, EditCtlsLibU::_IUpDownTextBoxEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragDrop, EditCtlsLibA::_IUpDownTextBoxEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data. If \c NULL, the cached data object is used. We use this when
	///            we call this method from other places than \c DragEnter.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragEnter,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragEnter,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragLeave,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::OLEDragLeave, Raise_OLEDragEnter, Raise_OLEDragMouseMove,
	///       Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragLeave,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::OLEDragLeave, Raise_OLEDragEnter, Raise_OLEDragMouseMove,
	///       Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragMouseMove,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::OLEDragMouseMove, Raise_OLEDragEnter, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OLEDragMouseMove,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::OLEDragMouseMove, Raise_OLEDragEnter, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OutOfMemory event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OutOfMemory, EditCtlsLibU::_IUpDownTextBoxEvents::OutOfMemory
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_OutOfMemory, EditCtlsLibA::_IUpDownTextBoxEvents::OutOfMemory
	/// \endif
	inline HRESULT Raise_OutOfMemory(void);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_RClick, EditCtlsLibU::_IUpDownTextBoxEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick,
	///       EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_RClick, EditCtlsLibA::_IUpDownTextBoxEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick,
	///       EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_RDblClick, EditCtlsLibU::_IUpDownTextBoxEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_RDblClick, EditCtlsLibA::_IUpDownTextBoxEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_RecreatedControlWindow,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_RecreatedControlWindow,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ResizedControlWindow,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ResizedControlWindow,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c TextChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_TextChanged, EditCtlsLibU::_IUpDownTextBoxEvents::TextChanged,
	///       Raise_ValueChanged
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_TextChanged, EditCtlsLibA::_IUpDownTextBoxEvents::TextChanged,
	///       Raise_ValueChanged
	/// \endif
	inline HRESULT Raise_TextChanged(void);
	/// \brief <em>Raises the \c TruncatedText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_TruncatedText,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::TruncatedText
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_TruncatedText,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::TruncatedText
	/// \endif
	inline HRESULT Raise_TruncatedText(void);
	/// \brief <em>Raises the \c ValueChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default event.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ValueChanged,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::ValueChanged, Raise_ValueChanging, Raise_TextChanged
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ValueChanged,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::ValueChanged, Raise_ValueChanging, Raise_TextChanged
	/// \endif
	inline HRESULT Raise_ValueChanged(void);
	/// \brief <em>Raises the \c ValueChanging event</em>
	///
	/// \param[in] currentValue The control's current value.
	/// \param[in] delta The value by which the control's value is about to change.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the value change; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ValueChanging,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::ValueChanging, Raise_ValueChanged, Raise_TextChanged
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_ValueChanging,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::ValueChanging, Raise_ValueChanged, Raise_TextChanged
	/// \endif
	inline HRESULT Raise_ValueChanging(LONG currentValue, LONG delta, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c WritingDirectionChanged event</em>
	///
	/// \param[in] newWritingDirection The contained edit control's new writing direction. Any of the values
	///            defined by the \c WritingDirectionConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to limitations of Microsoft Windows, this event is not raised if the writing direction
	///          is changed using the contained edit control's default context menu.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_WritingDirectionChanged,
	///       EditCtlsLibU::_IUpDownTextBoxEvents::WritingDirectionChanged,
	///       EditCtlsLibU::WritingDirectionConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_WritingDirectionChanged,
	///       EditCtlsLibA::_IUpDownTextBoxEvents::WritingDirectionChanged,
	///       EditCtlsLibA::WritingDirectionConstants
	/// \endif
	inline HRESULT Raise_WritingDirectionChanged(WritingDirectionConstants newWritingDirection);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_XClick, EditCtlsLibU::_IUpDownTextBoxEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_XClick, EditCtlsLibA::_IUpDownTextBoxEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_XDblClick, EditCtlsLibU::_IUpDownTextBoxEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IUpDownTextBoxEvents::Fire_XDblClick, EditCtlsLibA::_IUpDownTextBoxEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, get_DetectDoubleClicks,
	///       EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa RecreateEditWindow, RecreateUpDownWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name IME support
	///
	//@{
	/// \brief <em>Retrieves a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is requested.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa SetCurrentIMEContextMode, EditCtlsLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa SetCurrentIMEContextMode, EditCtlsLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetCurrentIMEContextMode(HWND hWnd);
	/// \brief <em>Sets a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is set.
	/// \param[in] IMEMode A constant out of the \c IMEModeConstants enumeration specifying the IME
	///            context mode to apply.
	///
	/// \if UNICODE
	///   \sa GetCurrentIMEContextMode, EditCtlsLibU::IMEModeConstants, put_IMEMode
	/// \else
	///   \sa GetCurrentIMEContextMode, EditCtlsLibA::IMEModeConstants, put_IMEMode
	/// \endif
	void SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode);
	/// \brief <em>Retrieves the control's effective IME context mode</em>
	///
	/// Retrieves the IME context mode that is set for the control after resolving recursive modes like
	/// \c imeInherit.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa EditCtlsLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa EditCtlsLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetEffectiveIMEMode(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Calculates the extended window style bits for the contained edit control</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetEditStyleBits, GetUpDownExStyleBits, SendConfigurationMessages, OnInitDialog
	DWORD GetEditExStyleBits(void);
	/// \brief <em>Calculates the window style bits for the contained edit control</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* and \c ES_* constants specifying the required window styles.
	///
	/// \sa GetEditExStyleBits, GetUpDownStyleBits, SendConfigurationMessages, OnInitDialog
	DWORD GetEditStyleBits(void);
	/// \brief <em>Calculates the extended window style bits for the contained edit control</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetUpDownStyleBits, GetEditExStyleBits, SendConfigurationMessages, OnInitDialog
	DWORD GetUpDownExStyleBits(void);
	/// \brief <em>Calculates the window style bits for the contained edit control</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* and \c ES_* constants specifying the required window styles.
	///
	/// \sa GetUpDownExStyleBits, GetEditStyleBits, SendConfigurationMessages, OnInitDialog
	DWORD GetUpDownStyleBits(void);
	/// \brief <em>Recreates the contained edit window</em>
	///
	/// Destroys and re-creates the contained edit window.
	///
	/// \sa RecreateUpDownWindow, RecreateControlWindow
	void RecreateEditWindow(void);
	/// \brief <em>Recreates the contained up down window</em>
	///
	/// Destroys and re-creates the contained up down window.
	///
	/// \sa RecreateEditWindow, RecreateControlWindow
	void RecreateUpDownWindow(void);
	/// \brief <em>Configures the control by sending messages to the contained controls</em>
	///
	/// Sends \c WM_*, \c EM_* and \c UDM_* messages to the contained control windows to make the control
	/// match the current property settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa SendEditConfigurationMessages, SendUpDownConfigurationMessages, GetEditExStyleBits,
	///     GetEditStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	/// \brief <em>Configures the control by sending messages to the contained edit control</em>
	///
	/// Sends \c WM_* and \c EM_* messages to the contained edit control window to make the control
	/// match the current property settings.
	///
	/// \sa SendConfigurationMessages, SendUpDownConfigurationMessages
	void SendEditConfigurationMessages(void);
	/// \brief <em>Configures the control by sending messages to the contained up down control</em>
	///
	/// Sends \c WM_* and \c UDM_* messages to the contained up down control window to make the control
	/// match the current property settings.
	///
	/// \sa SendConfigurationMessages, SendEditConfigurationMessages
	void SendUpDownConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \if UNICODE
	///   \sa EditCtlsLibU::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \else
	///   \sa EditCtlsLibA::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \endif
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);


	/// \brief <em>Holds constants and flags used with IME support</em>
	struct IMEFlags
	{
	protected:
		/// \brief <em>A table of IME modes to use for Chinese input language</em>
		///
		/// \sa GetIMECountryTable, japaneseIMETable, koreanIMETable
		static IMEModeConstants chineseIMETable[10];
		/// \brief <em>A table of IME modes to use for Japanese input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants japaneseIMETable[10];
		/// \brief <em>A table of IME modes to use for Korean input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants koreanIMETable[10];

	public:
		/// \brief <em>The handle of the default IME context</em>
		HIMC hDefaultIMC;

		IMEFlags()
		{
			hDefaultIMC = NULL;
		}

		/// \brief <em>Retrieves a table of IME modes to use for the current keyboard layout</em>
		///
		/// Retrieves a table of IME modes which can be used to map \c IME_CMODE_* constants to
		/// \c IMEModeConstants constants. The table depends on the current keyboard layout.
		///
		/// \param[in,out] table The IME mode table for the currently active keyboard layout.
		///
		/// \if UNICODE
		///   \sa EditCtlsLibU::IMEModeConstants, GetCurrentIMEContextMode
		/// \else
		///   \sa EditCtlsLibA::IMEModeConstants, GetCurrentIMEContextMode
		/// \endif
		static void GetIMECountryTable(IMEModeConstants table[10])
		{
			WORD languageID = LOWORD(GetKeyboardLayout(0));
			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				switch(languageID) {
					case MAKELANGID(LANG_JAPANESE, SUBLANG_DEFAULT):
						CopyMemory(table, japaneseIMETable, sizeof(japaneseIMETable));
						return;
						break;
					case MAKELANGID(LANG_KOREAN, SUBLANG_DEFAULT):
						CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
						return;
						break;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
				if(languageID == MAKELANGID(LANG_KOREAN, SUBLANG_SYS_DEFAULT)) {
					CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
					return;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if((languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SINGAPORE)) && (languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_MACAU))) {
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
		}
	} IMEFlags;

	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnEditSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnEditSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<UpDownTextBox> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(UpDownTextBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<UpDownTextBox> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<UpDownTextBox> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(UpDownTextBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<UpDownTextBox> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>Holds the \c AcceptNumbersOnly property's setting</em>
		///
		/// \sa get_AcceptNumbersOnly, put_AcceptNumbersOnly
		UINT acceptNumbersOnly : 1;
		/// \brief <em>Holds the \c AlwaysShowSelection property's setting</em>
		///
		/// \sa get_AlwaysShowSelection, put_AlwaysShowSelection
		UINT alwaysShowSelection : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c AutomaticallyCorrectValue property's setting</em>
		///
		/// \sa get_AutomaticallyCorrectValue, put_AutomaticallyCorrectValue
		UINT automaticallyCorrectValue : 1;
		/// \brief <em>Holds the \c AutomaticallySetText property's setting</em>
		///
		/// \sa get_AutomaticallySetText, put_AutomaticallySetText
		UINT automaticallySetText : 1;
		/// \brief <em>Holds the \c AutoScrolling property's setting</em>
		///
		/// \sa get_AutoScrolling, put_AutoScrolling
		AutoScrollingConstants autoScrolling;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c Base property's setting</em>
		///
		/// \sa get_Base, put_Base
		BaseConstants base;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c CancelIMECompositionOnSetFocus property's setting</em>
		///
		/// \sa get_CancelIMECompositionOnSetFocus, put_CancelIMECompositionOnSetFocus
		UINT cancelIMECompositionOnSetFocus : 1;
		/// \brief <em>Holds the \c CharacterConversion property's setting</em>
		///
		/// \sa get_CharacterConversion, put_CharacterConversion
		CharacterConversionConstants characterConversion;
		/// \brief <em>Holds the \c CompleteIMECompositionOnKillFocus property's setting</em>
		///
		/// \sa get_CompleteIMECompositionOnKillFocus, put_CompleteIMECompositionOnKillFocus
		UINT completeIMECompositionOnKillFocus : 1;
		/// \brief <em>Holds the \c CueBanner property's setting</em>
		///
		/// \sa get_CueBanner, put_CueBanner
		CComBSTR cueBanner;
		/// \brief <em>Holds the \c CurrentValue property's setting</em>
		///
		/// \sa get_CurrentValue, put_CurrentValue
		long currentValue;
		/// \brief <em>Holds the \c DetectDoubleClicks property's setting</em>
		///
		/// \sa get_DetectDoubleClicks, put_DetectDoubleClicks
		UINT detectDoubleClicks : 1;
		/// \brief <em>Holds the \c DisabledBackColor property's setting</em>
		///
		/// \sa get_DisabledBackColor, put_DisabledBackColor
		OLE_COLOR disabledBackColor;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DisabledForeColor property's setting</em>
		///
		/// \sa get_DisabledForeColor, put_DisabledForeColor
		OLE_COLOR disabledForeColor;
		/// \brief <em>Holds the \c DisplayCueBannerOnFocus property's setting</em>
		///
		/// \sa get_DisplayCueBannerOnFocus, put_DisplayCueBannerOnFocus
		UINT displayCueBannerOnFocus : 1;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c DoOEMConversion property's setting</em>
		///
		/// \sa get_DoOEMConversion, put_DoOEMConversion
		UINT doOEMConversion : 1;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c ForeColor property's setting</em>
		///
		/// \sa get_ForeColor, put_ForeColor
		OLE_COLOR foreColor;
		/// \brief <em>Holds the \c GroupDigits property's setting</em>
		///
		/// \sa get_GroupDigits, put_GroupDigits
		UINT groupDigits : 1;
		/// \brief <em>Holds the \c HAlignment property's setting</em>
		///
		/// \sa get_HAlignment, put_HAlignment
		HAlignmentConstants hAlignment;
		/// \brief <em>Holds the \c HotTracking property's setting</em>
		///
		/// \sa get_HotTracking, put_HotTracking
		UINT hotTracking : 1;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c IMEMode property's setting</em>
		///
		/// \sa get_IMEMode, put_IMEMode
		IMEModeConstants IMEMode;
		/// \brief <em>Holds the \c LeftMargin property's setting</em>
		///
		/// \sa get_LeftMargin, put_LeftMargin
		OLE_XSIZE_PIXELS leftMargin;
		/// \brief <em>Holds the \c Maximum property's setting</em>
		///
		/// \sa get_Maximum, put_Maximum
		long maximum;
		/// \brief <em>Holds the \c MaxTextLength property's setting</em>
		///
		/// \sa get_MaxTextLength, put_MaxTextLength
		long maxTextLength;
		/// \brief <em>Holds the \c Minimum property's setting</em>
		///
		/// \sa get_Minimum, put_Minimum
		long minimum;
		/// \brief <em>Holds the \c Modified property's setting</em>
		///
		/// \sa get_Modified, put_Modified
		UINT modified : 1;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c Orientation property's setting</em>
		///
		/// \sa get_Orientation, put_Orientation
		OrientationConstants orientation;
		/// \brief <em>Holds the \c ProcessArrowKeys property's setting</em>
		///
		/// \sa get_ProcessArrowKeys, put_ProcessArrowKeys
		UINT processArrowKeys : 1;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c ReadOnlyTextBox property's setting</em>
		///
		/// \sa get_ReadOnlyTextBox, put_ReadOnlyTextBox
		UINT readOnlyTextBox : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		UINT registerForOLEDragDrop : 1;
		/// \brief <em>Holds the \c RightMargin property's setting</em>
		///
		/// \sa get_RightMargin, put_RightMargin
		OLE_XSIZE_PIXELS rightMargin;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c Text property's setting</em>
		///
		/// \sa get_Text, put_Text
		CComBSTR text;
		/// \brief <em>Holds the \c UpDownPosition property's setting</em>
		///
		/// \sa get_UpDownPosition, put_UpDownPosition
		UpDownPositionConstants upDownPosition;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;
		/// \brief <em>Holds the \c WrapAtBoundaries property's setting</em>
		///
		/// \sa get_WrapAtBoundaries, put_WrapAtBoundaries
		UINT wrapAtBoundaries : 1;

		Properties()
		{
			ResetToDefaults();
		}

		~Properties()
		{
			Release();
		}

		/// \brief <em>Prepares the object for destruction</em>
		void Release(void)
		{
			font.Release();
			mouseIcon.Release();
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			acceptNumbersOnly = FALSE;
			alwaysShowSelection = FALSE;
			appearance = a3D;
			automaticallyCorrectValue = TRUE;
			automaticallySetText = TRUE;
			autoScrolling = asHorizontal;
			backColor = 0x80000000 | COLOR_WINDOW;
			base = bDecimal;
			borderStyle = bsNone;
			cancelIMECompositionOnSetFocus = FALSE;
			characterConversion = ccNone;
			completeIMECompositionOnKillFocus = FALSE;
			cueBanner = L"";
			currentValue = 0;
			detectDoubleClicks = FALSE;
			disabledBackColor = static_cast<OLE_COLOR>(-1);
			disabledEvents = static_cast<DisabledEventsConstants>(deBeforeDrawText | deClickEvents | deMouseEvents | deTextChangedEvents);
			disabledForeColor = static_cast<OLE_COLOR>(-1);
			displayCueBannerOnFocus = FALSE;
			dontRedraw = FALSE;
			doOEMConversion = FALSE;
			enabled = TRUE;
			foreColor = 0x80000000 | COLOR_WINDOWTEXT;
			groupDigits = TRUE;
			hAlignment = halRight;
			hotTracking = FALSE;
			hoverTime = -1;
			IMEMode = imeInherit;
			leftMargin = -1;
			maximum = 100;
			maxTextLength = -1;
			minimum = 0;
			modified = FALSE;
			mousePointer = mpDefault;
			orientation = oVertical;
			processArrowKeys = TRUE;
			processContextMenuKeys = TRUE;
			readOnlyTextBox = FALSE;
			registerForOLEDragDrop = FALSE;
			rightMargin = -1;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			supportOLEDragImages = TRUE;
			text = L"0";
			upDownPosition = udRightOfTextBox;
			useSystemFont = TRUE;
			wrapAtBoundaries = FALSE;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, the control has been activated by mouse and needs to be UI-activated by \c OnSetFocus</em>
		///
		/// ATL always UI-activates the control in \c OnMouseActivate. If the control is activated by mouse,
		/// \c WM_SETFOCUS is sent after \c WM_MOUSEACTIVATE, but Visual Basic 6 won't raise the \c Validate
		/// event if the control already is UI-activated when it receives the focus. Therefore we need to delay
		/// UI-activation.
		///
		/// \sa OnMouseActivate, OnEditSetFocus, OnUpDownSetFocus, OnKillFocus
		UINT uiActivationPending : 1;
		/// \brief <em>Holds whether the setting of the \c AutomaticallyCorrectValue property is ignored</em>
		///
		/// If greater than 0, \c OnSetRange ignores the \c AutomaticallyCorrectValue property. This is useful
		/// during control initialization. If 0, the property isn't ignored.
		///
		/// \sa OnSetRange, get_AutomaticallyCorrectValue
		int noAutoCorrection;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;

		Flags()
		{
			uiActivationPending = FALSE;
			noAutoCorrection = 0;
			usingThemes = FALSE;
		}
	} flags;


	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event already was raised</em>
		///
		/// \sa Raise_MouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event already was raised</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa Raise_MouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		POINT lastPosition;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the contained edit control's mouse status</em>
	MouseStatus mouseStatus_Edit;
	/// \brief <em>Holds the contained up down control's mouse status</em>
	MouseStatus mouseStatus_UpDown;

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pDropTargetHelper = NULL;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			drop_pDataObject = NULL;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			//
		}
	} dragDropStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
	} timers;

	/// \brief <em>The brush that the control's background is drawn with, if the control is themed and disabled</em>
	///
	/// \deprecated This member should be used on Windows XP only.
	CBrush themedBackBrush;
	/// \brief <em>Holds the \c wParam parameter of the next \c WM_CHAR message</em>
	///
	/// The VB6 runtime implements the message loop using \c TranslateMessageA. This destroys input of
	/// Unicode characters if IME is not used. For instance Hindi cannot be inputted although the control is
	/// a Unicode window. To workaround this problem, we check the message queue on each \c WM_KEYDOWN and
	/// \c WM_KEYUP message. If it contains a \c WM_CHAR message, we store its \c wParam parameter and use it
	/// instead of the broken one when handling this \c WM_CHAR message.
	///
	/// \sa OnEditChar, OnEditKeyDown, OnEditKeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	WPARAM cachedWParam;

private:
};     // UpDownTextBox

OBJECT_ENTRY_AUTO(__uuidof(UpDownTextBox), UpDownTextBox)