//////////////////////////////////////////////////////////////////////
/// \class TextBox
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c Edit</em>
///
/// This class superclasses \c Edit and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo \c IMEFlags is the name of a struct as well as a variable.
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa EditCtlsLibU::ITextBox
/// \else
///   \sa EditCtlsLibA::ITextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "EditCtlsU.h"
#else
	#include "EditCtlsA.h"
#endif
#include "_ITextBoxEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "CommonProperties.h"
#include "StringProperties.h"
#include "TargetOLEDataObject.h"
#include "SourceOLEDataObject.h"


class ATL_NO_VTABLE TextBox : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<ITextBox, &IID_ITextBox, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<ITextBox, &IID_ITextBox, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<TextBox>,
    public IOleControlImpl<TextBox>,
    public IOleObjectImpl<TextBox>,
    public IOleInPlaceActiveObjectImpl<TextBox>,
    public IViewObjectExImpl<TextBox>,
    public IOleInPlaceObjectWindowlessImpl<TextBox>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TextBox>,
    public Proxy_ITextBoxEvents<TextBox>,
    public IPersistStorageImpl<TextBox>,
    public IPersistPropertyBagImpl<TextBox>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<TextBox>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_TextBox, &__uuidof(_ITextBoxEvents), &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_TextBox, &__uuidof(_ITextBoxEvents), &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<TextBox>,
    public CComCoClass<TextBox, &CLSID_TextBox>,
    public CComControl<TextBox>,
    public IPerPropertyBrowsingImpl<TextBox>,
    public IDropTarget,
    public IDropSource,
    public IDropSourceNotify,
    public ICategorizeProperties,
    public ICreditsProvider
{
	friend class SourceOLEDataObject;

protected:
	/// \brief <em>A custom edit control message which retrieves the position of the insertion mark</em>
	///
	/// \sa EM_SETINSERTMARK, ECINSERTMARK, EM_GETINSERTMARKCOLOR
	#define EM_GETINSERTMARK (WM_USER + 2)
	/// \brief <em>A custom edit control message which sets the position of the insertion mark</em>
	///
	/// \sa EM_GETINSERTMARK, EM_SETINSERTMARKCOLOR
	#define EM_SETINSERTMARK (WM_USER + 3)
	/// \brief <em>A custom edit control message which retrieves the color of the insertion mark</em>
	///
	/// \sa EM_SETINSERTMARKCOLOR, EM_GETINSERTMARK
	#define EM_GETINSERTMARKCOLOR (WM_USER + 4)
	/// \brief <em>A custom edit control message which sets the color of the insertion mark</em>
	///
	/// \sa EM_GETINSERTMARKCOLOR, EM_SETINSERTMARK
	#define EM_SETINSERTMARKCOLOR (WM_USER + 5)

	/// \brief <em>A struct used with \c EM_GETINSERTMARK</em>
	///
	/// \sa EM_GETINSERTMARK
	typedef struct
	{
		/// \brief <em>The struct's size in bytes</em>
		UINT size;
		/// \brief <em>\c TRUE, if the insertion mark is displayed after the specified character; otherwise \c FALSE</em>
		BOOL afterChar;
		/// \brief <em>The zero-based index of the character at which the insertion mark is displayed</em>
		int characterIndex;
	} ECINSERTMARK, *LPECINSERTMARK;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	TextBox();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_IMEMODE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_TEXTBOX)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("TextBoxU"), WC_EDITW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("TextBoxA"), WC_EDITA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(TextBox)
			COM_INTERFACE_ENTRY(ITextBox)
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
			COM_INTERFACE_ENTRY(IDropSource)
			COM_INTERFACE_ENTRY(IDropSourceNotify)
		END_COM_MAP()

		BEGIN_PROP_MAP(TextBox)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AcceptNumbersOnly", DISPID_TXTBOX_ACCEPTNUMBERSONLY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AcceptTabKey", DISPID_TXTBOX_ACCEPTTABKEY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowDragDrop", DISPID_TXTBOX_ALLOWDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AlwaysShowSelection", DISPID_TXTBOX_ALWAYSSHOWSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_TXTBOX_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutoScrolling", DISPID_TXTBOX_AUTOSCROLLING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BackColor", DISPID_TXTBOX_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_TXTBOX_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CancelIMECompositionOnSetFocus", DISPID_TXTBOX_CANCELIMECOMPOSITIONONSETFOCUS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CharacterConversion", DISPID_TXTBOX_CHARACTERCONVERSION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CompleteIMECompositionOnKillFocus", DISPID_TXTBOX_COMPLETEIMECOMPOSITIONONKILLFOCUS, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("CueBanner", DISPID_TXTBOX_CUEBANNER, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("DisabledBackColor", DISPID_TXTBOX_DISABLEDBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_TXTBOX_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisabledForeColor", DISPID_TXTBOX_DISABLEDFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("DisplayCueBannerOnFocus", DISPID_TXTBOX_DISPLAYCUEBANNERONFOCUS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_TXTBOX_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DoOEMConversion", DISPID_TXTBOX_DOOEMCONVERSION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragScrollTimeBase", DISPID_TXTBOX_DRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_TXTBOX_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_TXTBOX_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_TXTBOX_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("FormattingRectangleHeight", DISPID_TXTBOX_FORMATTINGRECTANGLEHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FormattingRectangleLeft", DISPID_TXTBOX_FORMATTINGRECTANGLELEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FormattingRectangleTop", DISPID_TXTBOX_FORMATTINGRECTANGLETOP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FormattingRectangleWidth", DISPID_TXTBOX_FORMATTINGRECTANGLEWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HAlignment", DISPID_TXTBOX_HALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HoverTime", DISPID_TXTBOX_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IMEMode", DISPID_TXTBOX_IMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("InsertMarkColor", DISPID_TXTBOX_INSERTMARKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("InsertSoftLineBreaks", DISPID_TXTBOX_INSERTSOFTLINEBREAKS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("LeftMargin", DISPID_TXTBOX_LEFTMARGIN, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaxTextLength", DISPID_TXTBOX_MAXTEXTLENGTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Modified", DISPID_TXTBOX_MODIFIED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_TXTBOX_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_TXTBOX_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MultiLine", DISPID_TXTBOX_MULTILINE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("OLEDragImageStyle", DISPID_TXTBOX_OLEDRAGIMAGESTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("PasswordChar", DISPID_TXTBOX_PASSWORDCHAR, CLSID_NULL, VT_I2)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_TXTBOX_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ReadOnly", DISPID_TXTBOX_READONLY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_TXTBOX_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightMargin", DISPID_TXTBOX_RIGHTMARGIN, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_TXTBOX_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ScrollBars", DISPID_TXTBOX_SCROLLBARS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SelectedTextMouseIcon", DISPID_TXTBOX_SELECTEDTEXTMOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("SelectedTextMousePointer", DISPID_TXTBOX_SELECTEDTEXTMOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_TXTBOX_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("TabWidth", DISPID_TXTBOX_TABWIDTH, CLSID_NULL, VT_I4)
			//PROP_ENTRY_TYPE("Text", DISPID_TXTBOX_TEXT, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("UseCustomFormattingRectangle", DISPID_TXTBOX_USECUSTOMFORMATTINGRECTANGLE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UsePasswordChar", DISPID_TXTBOX_USEPASSWORDCHAR, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_TXTBOX_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(TextBox)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_ITextBoxEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(TextBox)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBkGnd)
			MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCode)
			MESSAGE_HANDLER(WM_HSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_IME_CHAR, OnIMEChar)
			MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnInputLangChange)
			MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTEXT, OnSetText)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_VSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM_CTLCOLOREDIT, OnReflectedCtlColorEdit)
			MESSAGE_HANDLER(OCM_CTLCOLORSTATIC, OnReflectedCtlColorStatic)

			MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)

			MESSAGE_HANDLER(EM_FMTLINES, OnFmtLines)
			MESSAGE_HANDLER(EM_GETINSERTMARK, OnGetInsertMark)
			MESSAGE_HANDLER(EM_GETINSERTMARKCOLOR, OnGetInsertMarkColor)
			MESSAGE_HANDLER(EM_SETCUEBANNER, OnSetCueBanner)
			MESSAGE_HANDLER(EM_SETINSERTMARK, OnSetInsertMark)
			MESSAGE_HANDLER(EM_SETINSERTMARKCOLOR, OnSetInsertMarkColor)
			MESSAGE_HANDLER(EM_SETTABSTOPS, OnSetTabStops)

			REFLECTED_COMMAND_CODE_HANDLER(EN_ALIGN_LTR_EC, OnReflectedAlign)
			REFLECTED_COMMAND_CODE_HANDLER(EN_ALIGN_RTL_EC, OnReflectedAlign)
			REFLECTED_COMMAND_CODE_HANDLER(EN_CHANGE, OnReflectedChange)
			REFLECTED_COMMAND_CODE_HANDLER(EN_ERRSPACE, OnReflectedErrSpace)
			REFLECTED_COMMAND_CODE_HANDLER(EN_HSCROLL, OnReflectedScroll)
			REFLECTED_COMMAND_CODE_HANDLER(EN_MAXTEXT, OnReflectedMaxText)
			REFLECTED_COMMAND_CODE_HANDLER(EN_SETFOCUS, OnReflectedSetFocus)
			REFLECTED_COMMAND_CODE_HANDLER(EN_UPDATE, OnReflectedUpdate)
			REFLECTED_COMMAND_CODE_HANDLER(EN_VSCROLL, OnReflectedScroll)

			CHAIN_MSG_MAP(CComControl<TextBox>)
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
	/// \name Implementation of ITextBox
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AcceptNumbersOnly property</em>
	///
	/// Retrieves whether the control accepts all kind of text or only numbers. If set to \c VARIANT_TRUE,
	/// only numbers, otherwise all text is accepted.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AcceptNumbersOnly, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_AcceptNumbersOnly(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AcceptNumbersOnly property</em>
	///
	/// Sets whether the control accepts all kind of text or only numbers. If set to \c VARIANT_TRUE,
	/// only numbers, otherwise all text is accepted.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AcceptNumbersOnly, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_AcceptNumbersOnly(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AcceptTabKey property</em>
	///
	/// Retrieves whether pressing the [TAB] key inserts a tabulator into the control. If set to
	/// \c VARIANT_TRUE, a tabulator is inserted; otherwise the keyboard focus is transfered to the next
	/// control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AcceptTabKey, get_TabStops, get_TabWidth, get_Text, Raise_KeyDown
	virtual HRESULT STDMETHODCALLTYPE get_AcceptTabKey(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AcceptTabKey property</em>
	///
	/// Sets whether pressing the [TAB] key inserts a tabulator into the control. If set to
	/// \c VARIANT_TRUE, a tabulator is inserted; otherwise the keyboard focus is transfered to the next
	/// control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AcceptTabKey, put_TabStops, put_TabWidth, put_Text, Raise_KeyDown
	virtual HRESULT STDMETHODCALLTYPE put_AcceptTabKey(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowDragDrop property</em>
	///
	/// Retrieves whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over selected text and then moving the
	/// mouse with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not available.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowDragDrop, get_RegisterForOLEDragDrop, get_DragScrollTimeBase, Raise_BeginDrag,
	///     Raise_BeginRDrag
	virtual HRESULT STDMETHODCALLTYPE get_AllowDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowDragDrop property</em>
	///
	/// Sets whether drag'n'drop mode can be entered. If set to \c VARIANT_TRUE, drag'n'drop mode
	/// can be entered by pressing the left or right mouse button over selected text and then moving the
	/// mouse with the button still pressed. If set to \c VARIANT_FALSE, drag'n'drop mode is not available.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowDragDrop, put_RegisterForOLEDragDrop, put_DragScrollTimeBase, Raise_BeginDrag,
	///     Raise_BeginRDrag
	virtual HRESULT STDMETHODCALLTYPE put_AllowDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AlwaysShowSelection property</em>
	///
	/// Retrieves whether the selected text will be highlighted even if the control doesn't have the focus.
	/// If set to \c VARIANT_TRUE, selected text is drawn as selected if the control does not have the focus;
	/// otherwise it's drawn as normal text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_AlwaysShowSelection, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_AlwaysShowSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AlwaysShowSelection property</em>
	///
	/// Sets whether the selected text will be highlighted even if the control doesn't have the focus.
	/// If set to \c VARIANT_TRUE, selected text is drawn as selected if the control does not have the focus;
	/// otherwise it's drawn as normal text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
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
	/// \brief <em>Retrieves the current setting of the \c AutoScrolling property</em>
	///
	/// Retrieves the directions into which the control scrolls automatically, if the caret reaches the
	/// borders of the control's client area. Any combination of the values defined by the
	/// \c AutoScrollingConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa put_AutoScrolling, get_ScrollBars, get_MultiLine, Raise_TruncatedText,
	///       EditCtlsLibU::AutoScrollingConstants
	/// \else
	///   \sa put_AutoScrolling, get_ScrollBars, get_MultiLine, Raise_TruncatedText,
	///       EditCtlsLibA::AutoScrollingConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AutoScrolling(AutoScrollingConstants* pValue);
	/// \brief <em>Sets the \c AutoScrolling property</em>
	///
	/// Sets the directions into which the control scrolls automatically, if the caret reaches the
	/// borders of the control's client area. Any combination of the values defined by the
	/// \c AutoScrollingConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_AutoScrolling, put_ScrollBars, put_MultiLine, Raise_TruncatedText,
	///       EditCtlsLibU::AutoScrollingConstants
	/// \else
	///   \sa get_AutoScrolling, put_ScrollBars, put_MultiLine, Raise_TruncatedText,
	///       EditCtlsLibA::AutoScrollingConstants
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
	/// Retrieves whether the control cancels the IME composition string when it receives the focus. If set
	/// to \c VARIANT_TRUE, the composition string is canceled; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CancelIMECompositionOnSetFocus, get_IMEMode, get_CompleteIMECompositionOnKillFocus
	virtual HRESULT STDMETHODCALLTYPE get_CancelIMECompositionOnSetFocus(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CancelIMECompositionOnSetFocus property</em>
	///
	/// Sets whether the control cancels the IME composition string when it receives the focus. If set
	/// to \c VARIANT_TRUE, the composition string is canceled; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CancelIMECompositionOnSetFocus, put_IMEMode, put_CompleteIMECompositionOnKillFocus
	virtual HRESULT STDMETHODCALLTYPE put_CancelIMECompositionOnSetFocus(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CharacterConversion property</em>
	///
	/// Retrieves the kind of conversion that is applied to characters that are typed into the control. Any
	/// of the values defined by the \c CharacterConversionConstants enumeration is valid.
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
	/// Sets the kind of conversion that is applied to characters that are typed into the control. Any
	/// of the values defined by the \c CharacterConversionConstants enumeration is valid.
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
	/// Retrieves whether the control completes the IME composition string when it loses the focus. If set to
	/// \c VARIANT_TRUE, the composition string is completed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CompleteIMECompositionOnKillFocus, get_IMEMode, get_CancelIMECompositionOnSetFocus
	virtual HRESULT STDMETHODCALLTYPE get_CompleteIMECompositionOnKillFocus(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CompleteIMECompositionOnKillFocus property</em>
	///
	/// Sets whether the control completes the IME composition string when it loses the focus. If set to
	/// \c VARIANT_TRUE, the composition string is completed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CompleteIMECompositionOnKillFocus, put_IMEMode, put_CancelIMECompositionOnSetFocus
	virtual HRESULT STDMETHODCALLTYPE put_CompleteIMECompositionOnKillFocus(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CueBanner property</em>
	///
	/// Retrieves the control's textual cue.
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
	/// Sets the control's textual cue.
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
	/// \brief <em>Retrieves the current setting of the \c DisabledBackColor property</em>
	///
	/// Retrieves the color used as the control's background color, if the control is read-only or
	/// disabled. If set to -1, the system's default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DisabledBackColor, get_Enabled, get_ReadOnly, get_DisabledForeColor, get_BackColor
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
	/// \sa get_DisabledBackColor, put_Enabled, put_ReadOnly, put_DisabledForeColor, put_BackColor
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
	/// \sa put_DisabledForeColor, get_Enabled, get_ReadOnly, get_DisabledBackColor, get_ForeColor
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
	/// \sa get_DisabledForeColor, put_Enabled, put_ReadOnly, put_DisabledBackColor, put_ForeColor
	virtual HRESULT STDMETHODCALLTYPE put_DisabledForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayCueBannerOnFocus property</em>
	///
	/// Retrieves whether the control's textual cue is displayed if the control has the keyboard focus.
	/// If set to \c VARIANT_TRUE, the textual cue is displayed if the control has the keyboard focus;
	/// otherwise not.
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
	/// Sets whether the control's textual cue is displayed if the control has the keyboard focus.
	/// If set to \c VARIANT_TRUE, the textual cue is displayed if the control has the keyboard focus;
	/// otherwise not.
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
	/// Retrieves whether the control's text is converted from the Windows character set to the OEM character
	/// set and then back to the Windows character set. Such a conversion ensures proper character conversion
	/// when the application calls the \c CharToOem function to convert a Windows string in the control to
	/// OEM characters. This property is most useful if the control contains file names that will be used on
	/// file systems that do not support Unicode.\n
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
	/// Sets whether the control's text is converted from the Windows character set to the OEM character
	/// set and then back to the Windows character set. Such a conversion ensures proper character conversion
	/// when the application calls the \c CharToOem function to convert a Windows string in the control to
	/// OEM characters. This property is most useful if the control contains file names that will be used on
	/// file systems that do not support Unicode.\n
	/// If set to \c VARIANT_TRUE, the conversion is performed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DoOEMConversion, put_CharacterConversion, put_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647473.aspx">CharToOem</a>
	virtual HRESULT STDMETHODCALLTYPE put_DoOEMConversion(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time, divided by 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragScrollTimeBase, get_AllowDragDrop, get_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c DragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time divided by 4 is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragScrollTimeBase, put_AllowDragDrop, put_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragScrollTimeBase(LONG newValue);
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
	/// Retrieves the zero-based index of the first visible character in a single-line control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FirstVisibleLine, GetLineFromChar, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FirstVisibleChar(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c FirstVisibleLine property</em>
	///
	/// Retrieves the zero-based index of the uppermost visible line in a multiline control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_LastVisibleLine, get_FirstVisibleChar, get_MultiLine, GetLineCount
	virtual HRESULT STDMETHODCALLTYPE get_FirstVisibleLine(LONG* pValue);
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
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleHeight, get_FormattingRectangleLeft, get_FormattingRectangleTop,
	///     get_FormattingRectangleWidth, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleHeight property</em>
	///
	/// Sets the height (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleHeight, put_FormattingRectangleLeft, put_FormattingRectangleTop,
	///     put_FormattingRectangleWidth, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleLeft property</em>
	///
	/// Retrieves the distance (in pixels) between the left borders of the control's formatting rectangle and
	/// its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleLeft, get_FormattingRectangleHeight, get_FormattingRectangleTop,
	///     get_FormattingRectangleWidth, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleLeft(OLE_XPOS_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleLeft property</em>
	///
	/// Retrieves the distance (in pixels) between the left borders of the control's formatting rectangle and
	/// its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleLeft, put_FormattingRectangleHeight, put_FormattingRectangleTop,
	///     put_FormattingRectangleWidth, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleLeft(OLE_XPOS_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleTop property</em>
	///
	/// Retrieves the distance (in pixels) between the upper borders of the control's formatting rectangle
	/// and its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleTop, get_FormattingRectangleHeight, get_FormattingRectangleLeft,
	///     get_FormattingRectangleWidth, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleTop(OLE_YPOS_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleTop property</em>
	///
	/// Sets the distance (in pixels) between the upper borders of the control's formatting rectangle
	/// and its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleTop, put_FormattingRectangleHeight, put_FormattingRectangleLeft,
	///     put_FormattingRectangleWidth, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleTop(OLE_YPOS_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleWidth property</em>
	///
	/// Retrieves the width (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleWidth, get_FormattingRectangleHeight, get_FormattingRectangleLeft,
	///     get_FormattingRectangleTop, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleWidth property</em>
	///
	/// Sets the width (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleWidth, put_FormattingRectangleHeight, put_FormattingRectangleLeft,
	///     put_FormattingRectangleTop, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c HAlignment property</em>
	///
	/// Retrieves the horizontal alignment of the control's content. Any of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention On Windows XP, changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa put_HAlignment, get_Text, EditCtlsLibU::HAlignmentConstants
	/// \else
	///   \sa put_HAlignment, get_Text, EditCtlsLibA::HAlignmentConstants
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
	/// \attention On Windows XP, changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_HAlignment, put_Text, EditCtlsLibU::HAlignmentConstants
	/// \else
	///   \sa get_HAlignment, put_Text, EditCtlsLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c hDragImageList property</em>
	///
	/// Retrieves the handle to the imagelist containing the drag image that is used during a
	/// drag'n'drop operation to visualize the dragged data.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, BeginDrag, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_hDragImageList(OLE_HANDLE* pValue);
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
	/// \sa Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
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
	///   \sa put_IMEMode, get_CancelIMECompositionOnSetFocus, get_CompleteIMECompositionOnKillFocus,
	///       EditCtlsLibU::IMEModeConstants
	/// \else
	///   \sa put_IMEMode, get_CancelIMECompositionOnSetFocus, get_CompleteIMECompositionOnKillFocus,
	///       EditCtlsLibA::IMEModeConstants
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
	///   \sa get_IMEMode, put_CancelIMECompositionOnSetFocus, put_CompleteIMECompositionOnKillFocus,
	///       EditCtlsLibU::IMEModeConstants
	/// \else
	///   \sa get_IMEMode, put_CancelIMECompositionOnSetFocus, put_CompleteIMECompositionOnKillFocus,
	///       EditCtlsLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IMEMode(IMEModeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c InsertMarkColor property</em>
	///
	/// Retrieves the color that is the control's insertion mark is drawn in.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_InsertMarkColor, GetInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE get_InsertMarkColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c InsertMarkColor property</em>
	///
	/// Sets the color that is the control's insertion mark is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_InsertMarkColor, SetInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE put_InsertMarkColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c InsertSoftLineBreaks property</em>
	///
	/// Retrieves whether the control inserts soft line-break characters at the end of lines that are broken
	/// because of wordwrapping. A soft line break consists of two carriage returns and a line feed. If set
	/// to \c VARIANT_TRUE, soft line breaks are inserted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_InsertSoftLineBreaks, get_Text, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_InsertSoftLineBreaks(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c InsertSoftLineBreaks property</em>
	///
	/// Sets whether the control inserts soft line-break characters at the end of lines that are broken
	/// because of wordwrapping. A soft line break consists of two carriage returns and a line feed. If set
	/// to \c VARIANT_TRUE, soft line breaks are inserted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_InsertSoftLineBreaks, put_Text, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_InsertSoftLineBreaks(VARIANT_BOOL newValue);
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
	/// \brief <em>Retrieves the current setting of the \c LastVisibleLine property</em>
	///
	/// Retrieves the zero-based index of the last visible line in a multiline control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FirstVisibleLine, get_MultiLine, GetLineCount
	virtual HRESULT STDMETHODCALLTYPE get_LastVisibleLine(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c LeftMargin property</em>
	///
	/// Retrieves the width (in pixels) of the control's left margin. If set to -1, a value, that depends on
	/// the control's font, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LeftMargin, get_RightMargin, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_LeftMargin(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c LeftMargin property</em>
	///
	/// Sets the width (in pixels) of the control's left margin. If set to -1, a value, that depends on
	/// the control's font, is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LeftMargin, put_RightMargin, put_Font
	virtual HRESULT STDMETHODCALLTYPE put_LeftMargin(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c LineHeight property</em>
	///
	/// Retrieves the height of each text line.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FirstVisibleLine, get_LastVisibleLine
	virtual HRESULT STDMETHODCALLTYPE get_LineHeight(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c MaxTextLength property</em>
	///
	/// Retrieves the maximum number of characters, that the user can type into the control. If set to -1,
	/// the system's default setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c Text property may exceed this limit.
	///
	/// \sa put_MaxTextLength, get_TextLength, get_Text, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE get_MaxTextLength(LONG* pValue);
	/// \brief <em>Sets the \c MaxTextLength property</em>
	///
	/// Sets the maximum number of characters, that the user can type into the control. If set to -1,
	/// the system's default setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c Text property may exceed this limit.
	///
	/// \sa get_MaxTextLength, get_TextLength, put_Text, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE put_MaxTextLength(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Modified property</em>
	///
	/// Retrieves a flag indicating whether the control's content has changed. A value of \c VARIANT_TRUE
	/// stands for changed content, a value of \c VARIANT_FALSE for unchanged content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Modified, get_Text, Raise_ContentChanged
	virtual HRESULT STDMETHODCALLTYPE get_Modified(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Modified property</em>
	///
	/// Sets a flag indicating whether the control's content has changed. A value of \c VARIANT_TRUE
	/// stands for changed content, a value of \c VARIANT_FALSE for unchanged content.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Modified, put_Text, Raise_ContentChanged
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
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, get_SelectedTextMouseIcon,
	///       EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, get_SelectedTextMouseIcon,
	///       EditCtlsLibA::MousePointerConstants,
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
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, put_SelectedTextMouseIcon,
	///       EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, put_SelectedTextMouseIcon,
	///       EditCtlsLibA::MousePointerConstants,
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
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, putref_SelectedTextMouseIcon,
	///       EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, putref_SelectedTextMouseIcon,
	///       EditCtlsLibA::MousePointerConstants,
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
	///   \sa put_MousePointer, get_MouseIcon, get_SelectedTextMousePointer,
	///       EditCtlsLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, get_SelectedTextMousePointer,
	///       EditCtlsLibA::MousePointerConstants
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
	///   \sa get_MousePointer, put_MouseIcon, put_SelectedTextMousePointer,
	///       EditCtlsLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, put_SelectedTextMousePointer,
	///       EditCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiLine property</em>
	///
	/// Retrieves whether the control processes carriage returns and displays the content on multiple
	/// lines. If set to \c VARIANT_TRUE, the content is displayed on multiple lines; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_MultiLine, get_Text, get_ScrollBars, GetLineCount, get_FirstVisibleLine, get_LastVisibleLine,
	///     get_HAlignment
	virtual HRESULT STDMETHODCALLTYPE get_MultiLine(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiLine property</em>
	///
	/// Sets whether the control processes carriage returns and displays the content on multiple
	/// lines. If set to \c VARIANT_TRUE, the content is displayed on multiple lines; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_MultiLine, put_Text, put_ScrollBars, put_HAlignment
	virtual HRESULT STDMETHODCALLTYPE put_MultiLine(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c OLEDragImageStyle property</em>
	///
	/// Retrieves the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       EditCtlsLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       EditCtlsLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue);
	/// \brief <em>Sets the \c OLEDragImageStyle property</em>
	///
	/// Sets the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       EditCtlsLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       EditCtlsLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OLEDragImageStyle(OLEDragImageStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c PasswordChar property</em>
	///
	/// Retrieves the code of the character, that is displayed instead of the real characters, if the
	/// \c UsePasswordChar property is set to \c VARIANT_TRUE. If set to 0, the system's default setting is
	/// used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_TRUE.
	///
	/// \sa put_PasswordChar, get_UsePasswordChar, get_Text, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_PasswordChar(SHORT* pValue);
	/// \brief <em>Sets the \c PasswordChar property</em>
	///
	/// Sets the code of the character, that is displayed instead of the real characters, if the
	/// \c UsePasswordChar property is set to \c VARIANT_TRUE. If set to 0, the system's default setting is
	/// used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_TRUE.
	///
	/// \sa get_PasswordChar, put_UsePasswordChar, put_Text, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_PasswordChar(SHORT newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu
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
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu
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
	/// \brief <em>Retrieves the current setting of the \c ReadOnly property</em>
	///
	/// Retrieves whether the control accepts user input, that would change the control's content. If set to
	/// \c VARIANT_FALSE, such user input is accepted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ReadOnly, get_Enabled, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ReadOnly(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ReadOnly property</em>
	///
	/// Sets whether the control accepts user input, that would change the control's content. If set to
	/// \c VARIANT_FALSE, such user input is accepted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ReadOnly, put_Enabled, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_ReadOnly(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RegisterForOLEDragDrop, get_AllowDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter
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
	/// \sa get_RegisterForOLEDragDrop, put_AllowDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RightMargin property</em>
	///
	/// Retrieves the width (in pixels) of the control's right margin. If set to -1, a value, that depends on
	/// the control's font, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RightMargin, get_LeftMargin, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_RightMargin(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c RightMargin property</em>
	///
	/// Sets the width (in pixels) of the control's right margin. If set to -1, a value, that depends on
	/// the control's font, is used.
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
	/// \attention On Windows XP, changing this property destroys and recreates the control window.
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
	/// \attention On Windows XP, changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, put_IMEMode, Raise_WritingDirectionChanged, EditCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, put_IMEMode, Raise_WritingDirectionChanged, EditCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ScrollBars property</em>
	///
	/// Retrieves the scrollbars to show. Any combination of the values defined by the \c ScrollBarsConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ScrollBars, get_AutoScrolling, get_MultiLine, Scroll, Raise_Scrolling,
	///       EditCtlsLibU::ScrollBarsConstants
	/// \else
	///   \sa put_ScrollBars, get_AutoScrolling, get_MultiLine, Scroll, Raise_Scrolling,
	///       EditCtlsLibA::ScrollBarsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ScrollBars(ScrollBarsConstants* pValue);
	/// \brief <em>Sets the \c ScrollBars property</em>
	///
	/// Sets the scrollbars to show. Any combination of the values defined by the \c ScrollBarsConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ScrollBars, put_AutoScrolling, put_MultiLine, Scroll, Raise_Scrolling,
	///       EditCtlsLibU::ScrollBarsConstants
	/// \else
	///   \sa get_ScrollBars, put_AutoScrolling, put_MultiLine, Scroll, Raise_Scrolling,
	///       EditCtlsLibA::ScrollBarsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ScrollBars(ScrollBarsConstants newValue);
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
	/// \brief <em>Retrieves the current setting of the \c SelectedTextMouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c SelectedTextMousePointer is set to
	/// \c mpCustom and the mouse cursor is located over selected text.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SelectedTextMouseIcon, putref_SelectedTextMouseIcon, get_SelectedTextMousePointer,
	///       get_MouseIcon, EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_SelectedTextMouseIcon, putref_SelectedTextMouseIcon, get_SelectedTextMousePointer,
	///       get_MouseIcon, EditCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectedTextMouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c SelectedTextMouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c SelectedTextMousePointer is set to
	/// \c mpCustom and the mouse cursor is located over selected text.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_SelectedTextMouseIcon, putref_SelectedTextMouseIcon, put_SelectedTextMousePointer,
	///       put_MouseIcon, EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_SelectedTextMouseIcon, putref_SelectedTextMouseIcon, put_SelectedTextMousePointer,
	///       put_MouseIcon, EditCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SelectedTextMouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c SelectedTextMouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c SelectedTextMousePointer is set to
	/// \c mpCustom and the mouse cursor is located over selected text.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SelectedTextMouseIcon, put_SelectedTextMouseIcon, put_SelectedTextMousePointer,
	///       putref_MouseIcon, EditCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_SelectedTextMouseIcon, put_SelectedTextMouseIcon, put_SelectedTextMousePointer,
	///       putref_MouseIcon, EditCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_SelectedTextMouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c SelectedTextMousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area over selected text. Any of the values defined by the \c MousePointerConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SelectedTextMousePointer, get_SelectedTextMouseIcon, get_MousePointer,
	///       EditCtlsLibU::MousePointerConstants
	/// \else
	///   \sa put_SelectedTextMousePointer, get_SelectedTextMouseIcon, get_MousePointer,
	///       EditCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectedTextMousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c SelectedTextMousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area over selected text. Any of the values defined by the \c MousePointerConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SelectedTextMousePointer, put_SelectedTextMouseIcon, put_MousePointer,
	///       EditCtlsLibU::MousePointerConstants
	/// \else
	///   \sa get_SelectedTextMousePointer, put_SelectedTextMouseIcon, put_MousePointer,
	///       EditCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SelectedTextMousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowDragImage property</em>
	///
	/// Retrieves whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible;
	/// otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowDragImage, get_hDragImageList, get_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_ShowDragImage(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowDragImage property</em>
	///
	/// Sets whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible; otherwise
	/// hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, get_hDragImageList, put_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_ShowDragImage(VARIANT_BOOL newValue);
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
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, get_ShowDragImage, get_OLEDragImageStyle,
	///     FinishOLEDragDrop,
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
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, put_ShowDragImage, put_OLEDragImageStyle,
	///     FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c TabStops property</em>
	///
	/// Retrieves the positions (in pixels) of the control's tab stops. The property expects a \c VARIANT
	/// containing an array of integer values, each specifying a tab stop's position.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_TabStops, get_AcceptTabKey, get_TabWidth
	virtual HRESULT STDMETHODCALLTYPE get_TabStops(VARIANT* pValue);
	/// \brief <em>Sets the \c TabStops property</em>
	///
	/// Sets the positions (in pixels) of the control's tab stops. The property expects a \c VARIANT
	/// containing an array of integer values, each specifying a tab stop's position.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_TabStops, put_AcceptTabKey, put_TabWidth
	virtual HRESULT STDMETHODCALLTYPE put_TabStops(VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c TabWidth property</em>
	///
	/// Retrieves the distance (in pixels) between 2 tab stops. If set to -1, the system's default value is
	/// used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.\n
	///          This property is ignored, if the \c TabStops property is not set to \c Empty.
	///
	/// \sa put_TabWidth, get_AcceptTabKey, get_TabStops
	virtual HRESULT STDMETHODCALLTYPE get_TabWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c TabWidth property</em>
	///
	/// Sets the distance (in pixels) between 2 tab stops. If set to -1, the system's default value is
	/// used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.\n
	///          This property is ignored, if the \c TabStops property is not set to \c Empty.
	///
	/// \sa get_TabWidth, put_AcceptTabKey, put_TabStops
	virtual HRESULT STDMETHODCALLTYPE put_TabWidth(OLE_XSIZE_PIXELS newValue);
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
	/// Retrieves the control's content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa put_Text, get_TextLength, get_MaxTextLength, GetLine, get_AcceptNumbersOnly, get_PasswordChar,
	///     get_CueBanner, get_HAlignment, get_MultiLine, get_ForeColor, get_Font, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the control's content.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa get_Text, get_TextLength, put_MaxTextLength, GetLine, put_AcceptNumbersOnly, put_PasswordChar,
	///     put_CueBanner, put_HAlignment, put_MultiLine, put_ForeColor, put_Font, Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c TextLength property</em>
	///
	/// Retrieves the length of the text specified by the \c Text property.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MaxTextLength, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_TextLength(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c UseCustomFormattingRectangle property</em>
	///
	/// Retrieves whether the control uses the formatting rectangle defined by the \c FormattingRectangle*
	/// properties.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for
	/// formatting the text displayed in the window rectangle. When an edit control is first displayed, the
	/// two rectangles are identical on the screen. An application can make the formatting rectangle larger
	/// than the window rectangle (thereby limiting the visibility of the control's text) or smaller than
	/// the window rectangle (thereby creating extra white space around the text).\n
	/// If this property is set to \c VARIANT_FALSE, the formatting rectangle is set to its default values.
	/// Otherwise it's defined by the \c FormattingRectangle* properties.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_UseCustomFormattingRectangle, get_FormattingRectangleHeight, get_FormattingRectangleLeft,
	///     get_FormattingRectangleTop, get_FormattingRectangleWidth, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_UseCustomFormattingRectangle(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseCustomFormattingRectangle property</em>
	///
	/// Sets whether the control uses the formatting rectangle defined by the \c FormattingRectangle*
	/// properties.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for
	/// formatting the text displayed in the window rectangle. When an edit control is first displayed, the
	/// two rectangles are identical on the screen. An application can make the formatting rectangle larger
	/// than the window rectangle (thereby limiting the visibility of the control's text) or smaller than
	/// the window rectangle (thereby creating extra white space around the text).\n
	/// If this property is set to \c VARIANT_FALSE, the formatting rectangle is set to its default values.
	/// Otherwise it's defined by the \c FormattingRectangle* properties.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_UseCustomFormattingRectangle, put_FormattingRectangleHeight, put_FormattingRectangleLeft,
	///     put_FormattingRectangleTop, put_FormattingRectangleWidth, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_UseCustomFormattingRectangle(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UsePasswordChar property</em>
	///
	/// Retrieves whether the control hides user input by (visually) replacing each character with the
	/// character specified by the \c PasswordChar property. If set to \c VARIANT_TRUE, user input is
	/// (visually) replaced; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_TRUE.
	///
	/// \sa put_UsePasswordChar, get_PasswordChar, get_Text, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_UsePasswordChar(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UsePasswordChar property</em>
	///
	/// Sets whether the control hides user input by (visually) replacing each character with the
	/// character specified by the \c PasswordChar property. If set to \c VARIANT_TRUE, user input is
	/// (visually) replaced; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_TRUE.
	///
	/// \sa get_UsePasswordChar, put_PasswordChar, put_Text, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_UsePasswordChar(VARIANT_BOOL newValue);
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
	/// Retrieves the function that is responsible to tell the control where a word starts and where it ends.
	/// This property takes the address of a function having the following signature:\n
	/// \code
	///   int CALLBACK FindWorkBreak(LPTSTR pText, int startPosition, int textLength, int flags);
	/// \endcode
	/// The \c pText argument is a pointer to the control's text.\n
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
	/// Sets the function that is responsible to tell the control where a word starts and where it ends.
	/// This property takes the address of a function having the following signature:\n
	/// \code
	///   int CALLBACK FindWorkBreak(LPTSTR pText, int startPosition, int textLength, int flags);
	/// \endcode
	/// The \c pText argument is a pointer to the control's text.\n
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

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Appends the specified text to the end of the current text</em>
	///
	/// \param[in] text The text to append.
	/// \param[in] setCaretToEnd If \c VARIANT_TRUE, the caret is automatically moved to the end of the text;
	///            otherwise not.
	/// \param[in] scrollToCaret If \c VARIANT_TRUE, the control is automatically scrolled to the caret;
	///            otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text
	virtual HRESULT STDMETHODCALLTYPE AppendText(BSTR text, VARIANT_BOOL setCaretToEnd = VARIANT_FALSE, VARIANT_BOOL scrollToCaret = VARIANT_TRUE);
	/// \brief <em>Enters drag'n'drop mode</em>
	///
	/// \param[in] draggedTextFirstChar The zero-based index of the first character of the text to drag.
	/// \param[in] draggedTextLastChar The zero-based index of the last character of the text to drag.
	/// \param[in] hDragImageList The imagelist containing the drag image that shall be used to
	///            visualize the drag'n'drop operation. If -1, the method creates the drag image itself;
	///            if \c NULL, no drag image is used.
	/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
	///                drag image's upper-left corner. If the \c hDragImageList parameter is set to -1 or
	///                \c NULL, this parameter is ignored. This parameter will be changed to the value that
	///                finally was used by the method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OLEDrag, GetDraggedTextRange, EndDrag, get_hDragImageList, Raise_BeginDrag, Raise_BeginRDrag,
	///     CreateDragImage
	virtual HRESULT STDMETHODCALLTYPE BeginDrag(LONG draggedTextFirstChar, LONG draggedTextLastChar, OLE_HANDLE hDragImageList = NULL, OLE_XPOS_PIXELS* pXHotSpot = NULL, OLE_YPOS_PIXELS* pYHotSpot = NULL);
	/// \brief <em>Determines whether there are any actions in the control's undo queue</em>
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
	/// \param[in] characterIndex The zero-based index of the character within the control, for which to
	///            retrieve the position. If the character is a line delimiter, the returned coordinates
	///            indicate a point just beyond the last visible character in the line. If the specified
	///            index is greater than the index of the last character in the control, the function fails.
	/// \param[out] pX The x-coordinate (in pixels) of the character relative to the control's upper-left
	///             corner.
	/// \param[out] pY The y-coordinate (in pixels) of the character relative to the control's upper-left
	///             corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PositionToCharIndex
	virtual HRESULT STDMETHODCALLTYPE CharIndexToPosition(LONG characterIndex, OLE_XPOS_PIXELS* pX = NULL, OLE_YPOS_PIXELS* pY = NULL);
	/// \brief <em>Clears the control's undo queue</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanUndo, Undo
	virtual HRESULT STDMETHODCALLTYPE EmptyUndoBuffer(void);
	/// \brief <em>Exits drag'n'drop mode</em>
	///
	/// \param[in] abort If \c VARIANT_TRUE, the drag'n'drop operation will be handled as aborted;
	///            otherwise it will be handled as a drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDraggedTextRange, BeginDrag, Raise_AbortedDrag, Raise_Drop
	virtual HRESULT STDMETHODCALLTYPE EndDrag(VARIANT_BOOL abort);
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
	/// \brief <em>Proposes a position for the control's insertion mark</em>
	///
	/// Retrieves the insertion mark position that is closest to the specified point.
	///
	/// \param[in] x The x-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified character. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] pCharacterIndex Receives the zero-based index of the character, at which the insertion
	///             mark should be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, EditCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, EditCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, LONG* pCharacterIndex);
	/// \brief <em>Retrieves the dragged text's start and end</em>
	///
	/// Retrieves the zero-based character indices of the dragged text's start and end.
	///
	/// \param[out] pDraggedTextFirstChar The zero-based index of the character at which the dragged text
	///             starts.
	/// \param[out] pDraggedTextLastChar The zero-based index of the last character of the dragged text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa BeginDrag, OLEDrag
	virtual HRESULT STDMETHODCALLTYPE GetDraggedTextRange(LONG* pDraggedTextFirstChar = NULL, LONG* pDraggedTextLastChar = NULL);
	/// \brief <em>Retrieves the zero-based index of the first character of the specified line</em>
	///
	/// \param[in] lineIndex The zero-based index of the line to retrieve the first character for. If set
	///            to -1, the index of the line containing the caret is used.
	/// \param[out] pValue The zero-based index of the first character of the line. -1 if the specified line
	///             index is greater than the total number of lines.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetLineFromChar, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE GetFirstCharOfLine(LONG lineIndex, LONG* pValue);
	/// \brief <em>Retrieves the position of the control's insertion mark</em>
	///
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified character. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] pCharacterIndex Receives the zero-based index of the character, at which the insertion
	///             mark is displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition, EditCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition, EditCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, LONG* pCharacterIndex);
	/// \brief <em>Retrieves the text of the specified line</em>
	///
	/// \param[in] lineIndex The zero-based index of the line to retrieve the text for.
	/// \param[out] pValue The line's text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetLineLength, get_MultiLine, get_Text
	virtual HRESULT STDMETHODCALLTYPE GetLine(LONG lineIndex, BSTR* pValue);
	/// \brief <em>Retrieves the number of lines in the control</em>
	///
	/// \param[out] pValue The number of lines.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstVisibleLine, get_LastVisibleLine, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE GetLineCount(LONG* pValue);
	/// \brief <em>Retrieves the zero-based index of the line that contains the specified character</em>
	///
	/// \param[in] characterIndex The zero-based index of the character within the control. If set to
	///            -1, the index of the character at which the selection begins, is used. If there's no
	///            selection, the index of the character next to the caret is used.
	/// \param[out] pValue The zero-based index of the line containing the character.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstVisibleChar, GetFirstCharOfLine, get_MultiLine, GetSelection
	virtual HRESULT STDMETHODCALLTYPE GetLineFromChar(LONG characterIndex, LONG* pValue);
	/// \brief <em>Retrieves the number of characters in the specified line</em>
	///
	/// \param[in] lineIndex The zero-based index of the line to retrieve the length for. If set to -1, the
	///            number of unselected characters on lines containing selected characters is retrieved.
	///            E. g. if the selection extended from the fourth character of one line through the eighth
	///            character from the end of the next line, the return value would be 10 (three characters on
	///            the first line and seven on the next).
	/// \param[out] pValue The number of characters in the line.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetLine, MultiLine
	virtual HRESULT STDMETHODCALLTYPE GetLineLength(LONG lineIndex, LONG* pValue);
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
	/// \brief <em>Hides any balloon tips associated with the control</em>
	///
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa ShowBalloonTip
	virtual HRESULT STDMETHODCALLTYPE HideBalloonTip(VARIANT_BOOL* pSucceeded);
	/// \brief <em>Determines whether the specified line's text is entirely visible or truncated</em>
	///
	/// \param[in] lineIndex The zero-based index of the line to check.
	/// \param[out] pValue \c VARIANT_TRUE if not all of the specified line's text is visible; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetLine
	virtual HRESULT STDMETHODCALLTYPE IsTextTruncated(LONG lineIndex = 0, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Enters OLE drag'n'drop mode</em>
	///
	/// \param[in] pDataObject A pointer to the \c IDataObject implementation to use during OLE
	///            drag'n'drop. If not specified, the control's own implementation is used.
	/// \param[in] supportedEffects A bit field defining all drop effects the client wants to support.
	///            Any combination of the values defined by the \c OLEDropEffectConstants enumeration
	///            (except \c odeScroll) is valid.
	/// \param[in] hWndToAskForDragImage The handle of the window, that is awaiting the
	///            \c DI_GETDRAGIMAGE message to specify the drag image to use. If -1, the method
	///            creates the drag image itself. If \c SupportOLEDragImages is set to \c VARIANT_FALSE,
	///            no drag image is used.
	/// \param[in] draggedTextFirstChar The zero-based index of the first character of the text to drag. This
	///            parameter is used to generate the drag image, if \c hWndToAskForDragImage is set to -1.
	/// \param[in] draggedTextLastChar The zero-based index of the last character of the text to drag. This
	///            parameter is used to generate the drag image, if \c hWndToAskForDragImage is set to -1.
	/// \param[in] itemCountToDisplay The number to display in the item count label of Aero drag images.
	///            If set to 0 or 1, no item count label is displayed. If set to any value larger than 1,
	///            this value is displayed in the item count label.
	/// \param[out] pPerformedEffects The performed drop effect. Any of the values defined by the
	///             \c OLEDropEffectConstants enumeration (except \c odeScroll) is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa BeginDrag, Raise_BeginDrag, Raise_BeginRDrag, Raise_OLEStartDrag, get_SupportOLEDragImages,
	///       get_OLEDragImageStyle, EditCtlsLibU::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \else
	///   \sa BeginDrag, Raise_BeginDrag, Raise_BeginRDrag, Raise_OLEStartDrag, get_SupportOLEDragImages,
	///       get_OLEDragImageStyle, EditCtlsLibA::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE OLEDrag(LONG* pDataObject = NULL, OLEDropEffectConstants supportedEffects = odeCopyOrMove, OLE_HANDLE hWndToAskForDragImage = -1, LONG draggedTextFirstChar = -1, LONG draggedTextLastChar = -1, LONG itemCountToDisplay = 0, OLEDropEffectConstants* pPerformedEffects = NULL);
	/// \brief <em>Retrieves the character closest to the specified position</em>
	///
	/// Retrieves the zero-based index of the character nearest the specified position.
	///
	/// \param[in] x The x-coordinate (in pixels) of the position to retrieve the nearest character for. It
	///            is relative to the control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the position to retrieve the nearest character for. It
	///            is relative to the control's upper-left corner.
	/// \param[out] pCharacterIndex The zero-based index of the character within the control, that is
	///             nearest to the specified position. If the specified point is beyond the last character
	///             in the control, this value indicates the last character in the control. The index
	///             indicates the line delimiter if the specified point is beyond the last visible
	///             character in a line.
	/// \param[out] pLineIndex The zero-based index of the line, that contains the character specified by
	///             the \c pCharacterIndex parameter.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If a point outside the bounds of the control is passed, the function fails.
	///
	/// \sa CharIndexToPosition
	virtual HRESULT STDMETHODCALLTYPE PositionToCharIndex(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG* pCharacterIndex = NULL, LONG* pLineIndex = NULL);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Replaces the currently selected text</em>
	///
	/// Replaces the control's currently selected text.
	///
	/// \param[in] replacementText The text that replaces the currently selected text.
	/// \param[in] undoable If \c VARIANT_TRUE, this action is inserted into the control's undo queue;
	///            otherwise not.
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
	/// \brief <em>Scrolls the control</em>
	///
	/// \param[in] axis The axis which is to be scrolled. Any combination of the values defined by the
	///            \c ScrollAxisConstants enumeration is valid.
	/// \param[in] directionAndIntensity The intensity and direction of the action. Any of the values defined
	///            by the \c ScrollDirectionConstants enumeration is valid.
	/// \param[in] linesToScrollVertically The number of lines to scroll vertically. This parameter is
	///            ignored, if \c directionAndIntensity is not set to \c sdCustom.
	/// \param[in] charactersToScrollHorizontally The number of characters to scroll horizontally. This
	///            parameter is ignored, if \c directionAndIntensity is not set to \c sdCustom.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method has no effect if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \if UNICODE
	///   \sa ScrollCaretIntoView, get_ScrollBars, get_AutoScrolling, get_MultiLine, Raise_Scrolling,
	///       EditCtlsLibU::ScrollAxisConstants, EditCtlsLibU::ScrollDirectionConstants
	/// \else
	///   \sa ScrollCaretIntoView, get_ScrollBars, get_AutoScrolling, get_MultiLine, Raise_Scrolling,
	///       EditCtlsLibA::ScrollAxisConstants, EditCtlsLibA::ScrollDirectionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Scroll(ScrollAxisConstants axis, ScrollDirectionConstants directionAndIntensity, LONG linesToScrollVertically = 0, LONG charactersToScrollHorizontally = 0);
	/// \brief <em>Scrolls the control so that the caret is visible</em>
	///
	/// Ensures that the control's caret is visible by scrolling the control if necessary.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Scroll, get_ScrollBars, get_AutoScrolling, get_MultiLine, Raise_Scrolling
	virtual HRESULT STDMETHODCALLTYPE ScrollCaretIntoView(void);
	/// \brief <em>Sets the position of the control's insertion mark</em>
	///
	/// \param[in] relativePosition The insertion mark's position relative to the specified character. Any of
	///            the values defined by the \c InsertMarkPositionConstants enumeration is valid.
	/// \param[in] characterIndex The zero-based index of the character, at which the insertion mark will be
	///            displayed. If set to -1, the insertion mark will be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor, get_AllowDragDrop,
	///       get_RegisterForOLEDragDrop, EditCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor, get_AllowDragDrop,
	///       get_RegisterForOLEDragDrop, EditCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, LONG characterIndex);
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
	/// \remarks To select all text in the control, set \c selectionStart to 0 and \c selectionEnd to -1.
	///
	/// \sa GetSelection, ReplaceSelectedText, get_SelectedText
	virtual HRESULT STDMETHODCALLTYPE SetSelection(LONG selectionStart, LONG selectionEnd);
	/// \brief <em>Displays a balloon tip associated with the control</em>
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
	/// \brief <em>Undoes the last action in the control's undo queue</em>
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
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
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
	/// \sa OnCreate, OnDestroy
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

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

	//////////////////////////////////////////////////////////////////////
	/// \name Drag image creation
	///
	//@{
	/// \brief <em>Creates a legacy drag image for the specified text</em>
	///
	/// Creates a drag image for the specified text in the style of Windows versions prior to Vista. The
	/// drag image is added to an imagelist which is returned.
	///
	/// \param[in] firstChar The zero-based index of the first character of the text for which to create the
	///            drag image.
	/// \param[in] lastChar The zero-based index of the last character of the text for which to create the
	///            drag image.
	/// \param[out] pUpperLeftPoint Receives the coordinates (in pixels) of the drag image's upper-left
	///             corner relative to the control's upper-left corner.
	/// \param[out] pBoundingRectangle Receives the drag image's bounding rectangle (in pixels) relative to
	///             the control's upper-left corner.
	///
	/// \return An imagelist containing the drag image.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa CreateLegacyOLEDragImage
	HIMAGELIST CreateLegacyDragImage(int firstChar, int lastChar, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle);
	/// \brief <em>Creates a legacy OLE drag image for the specified text</em>
	///
	/// Creates an OLE drag image for the specified text in the style of Windows versions prior to Vista.
	///
	/// \param[in] firstChar The zero-based index of the first character of the text for which to create the
	///            drag image.
	/// \param[in] lastChar The zero-based index of the last character of the text for which to create the
	///            drag image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateVistaOLEDragImage, CreateLegacyDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateLegacyOLEDragImage(int firstChar, int lastChar, __in LPSHDRAGIMAGE pDragImage);
	//@}
	//////////////////////////////////////////////////////////////////////

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
	/// \name Implementation of IDropSource
	///
	//@{
	/// \brief <em>Notifies the source of the current drop effect during OLE drag'n'drop</em>
	///
	/// This method is called frequently by the \c DoDragDrop function to notify the source of the
	/// last drop effect that the target has chosen. The source should set an appropriate mouse cursor.
	///
	/// \param[in] effect The drop effect chosen by the target. Any of the values defined by the
	///            \c DROPEFFECT enumeration is valid.
	///
	/// \return \c S_OK if the method has set a custom mouse cursor; \c DRAGDROP_S_USEDEFAULTCURSORS to
	///         use default mouse cursors; or an error code otherwise.
	///
	/// \sa QueryContinueDrag, Raise_OLEGiveFeedback,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693723.aspx">IDropSource::GiveFeedback</a>
	virtual HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD effect);
	/// \brief <em>Determines whether a drag'n'drop operation should be continued, canceled or completed</em>
	///
	/// This method is called by the \c DoDragDrop function to determine whether a drag'n'drop
	/// operation should be continued, canceled or completed.
	///
	/// \param[in] pressedEscape Indicates whether the user has pressed the \c ESC key since the
	///            previous call of this method. If \c TRUE, the key has been pressed; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	///
	/// \return \c S_OK if the drag'n'drop operation should continue; \c DRAGDROP_S_DROP if it should
	///         be completed; \c DRAGDROP_S_CANCEL if it should be canceled; or an error code otherwise.
	///
	/// \sa GiveFeedback, Raise_OLEQueryContinueDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690076.aspx">IDropSource::QueryContinueDrag</a>
	virtual HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL pressedEscape, DWORD keyState);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSourceNotify
	///
	//@{
	/// \brief <em>Notifies the source that the user drags the mouse cursor into a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor into a potential drop target window.
	///
	/// \param[in] hWndTarget The potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragLeaveTarget, Raise_OLEDragEnterPotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragEnterTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnterTarget(HWND hWndTarget);
	/// \brief <em>Notifies the source that the user drags the mouse cursor out of a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor out of a potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnterTarget, Raise_OLEDragLeavePotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragLeaveTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeaveTarget(void);
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
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, OnIMEChar, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_ERASEBKGND handler</em>
	///
	/// Will be called if the control's window background must be erased.
	/// We use this handler to set the custom formatting rectangle again, because it is overridden by
	/// Windows.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms648055.aspx">WM_ERASEBKGND</a>
	LRESULT OnEraseBkGnd(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_GETDLGCODE handler</em>
	///
	/// Will be called if the system needs to know which input messages the control wants to handle.
	/// We use this handler for the \c AcceptTabKey property.
	///
	/// \sa OnKeyDown, OnKeyUp, get_AcceptTabKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645425.aspx">WM_GETDLGCODE</a>
	LRESULT OnGetDlgCode(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_IME_CHAR handler</em>
	///
	/// Sent by the input method editor (IME) when a character has been composited.
	/// We use this handler to make more IME implementations (namely the one for emojis) work with ANSI
	/// applications like all VB6 applications.
	///
	/// \sa OnChar,
	///     <a href="https://docs.microsoft.com/en-us/windows/win32/intl/wm-ime-char">WM_IME_CHAR</a>
	LRESULT OnIMEChar(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_INPUTLANGCHANGE handler</em>
	///
	/// Will be called after an application's input language has been changed.
	/// We use this handler to update the IME mode of the control.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632629.aspx">WM_INPUTLANGCHANGE</a>
	LRESULT OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KILLFOCUS handler</em>
	///
	/// Will be called after the control gained the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnSetFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646282.aspx">WM_KILLFOCUS</a>
	LRESULT OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnMButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnSetFocus, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, OnMouseWheel, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnMouseMove, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT and \c WM_PRINTCLIENT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl. This makes Vista's graphic
	/// effects work.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534913.aspx">WM_PRINTCLIENT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_*SCROLL handler</em>
	///
	/// Will be called if the control shall be scrolled.
	/// We use this handler to raise the \c Scrolling event for the case that the scrollbar thumb is dragged,
	/// because in this case no \c EN_HSCROLL and \c EN_VSCROLL notification respectively is sent.
	///
	/// \sa OnReflectedScroll, Raise_Scrolling,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651283.aspx">WM_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651284.aspx">WM_VSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672113.aspx">EN_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672118.aspx">EN_VSCROLL</a>
	LRESULT OnScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer, get_SelectedTextMousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called after the control gained the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTEXT handler</em>
	///
	/// Will be called if the control's text shall be changed.
	/// We use this handler to raise the \c TextChanged event for multiline mode, because in this mode,
	/// \c WM_SETTEXT doesn't cause an \c EN_CHANGE notification.
	///
	/// \sa OnReflectedChange, get_Text, Raise_TextChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632644.aspx">WM_SETTEXT</a>
	LRESULT OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
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
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CTLCOLOREDIT handler</em>
	///
	/// Will be called if the control asks its parent window to configure the specified device context for
	/// drawing the control, i. e. to setup the colors and brushes.
	/// We use this handler for the \c BackColor and \c ForeColor properties.
	///
	/// \sa get_BackColor, get_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672119.aspx">WM_CTLCOLOREDIT</a>
	LRESULT OnReflectedCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CTLCOLORSTATIC handler</em>
	///
	/// Will be called if the control asks its parent window to configure the specified device context for
	/// drawing the control, i. e. to setup the colors and brushes.
	/// We use this handler for proper theming support and for the \c Enabled and \c ReadOnly properties.
	///
	/// \sa get_DisabledBackColor, get_DisabledForeColor, get_Enabled, get_ReadOnly,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651165.aspx">WM_CTLCOLORSTATIC</a>
	LRESULT OnReflectedCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DI_GETDRAGIMAGE handler</em>
	///
	/// Will be called during OLE drag'n'drop if the control is queried for a drag image.
	///
	/// \sa OLEDrag, CreateLegacyOLEDragImage, CreateVistaOLEDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	LRESULT OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c EM_FMTLINES handler</em>
	///
	/// Will be called if the control's soft line break settings shall be changed.
	/// We use this handler to synchronize our settings.
	///
	/// \sa get_InsertSoftLineBreaks,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672068.aspx">EM_FMTLINES</a>
	LRESULT OnFmtLines(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c EM_GETINSERTMARK handler</em>
	///
	/// Will be called if the position of the control's insertion mark shall be retrieved.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnSetInsertMark, insertMark, GetInsertMarkPosition, OnGetInsertMarkColor, EM_GETINSERTMARK
	LRESULT OnGetInsertMark(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c EM_GETINSERTMARKCOLOR handler</em>
	///
	/// Will be called if the color of the control's insertion mark shall be retrieved.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnSetInsertMarkColor, insertMark, OnGetInsertMark, EM_GETINSERTMARKCOLOR
	LRESULT OnGetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c EM_SETCUEBANNER handler</em>
	///
	/// Will be called if the control's textual cue shall be set.
	/// We use this handler to synchronize our settings.
	///
	/// \sa get_DisplayCueBannerOnFocus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672093.aspx">EM_SETCUEBANNER</a>
	LRESULT OnSetCueBanner(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c EM_SETINSERTMARK handler</em>
	///
	/// Will be called if the control's insertion mark shall be repositioned.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnGetInsertMark, insertMark, SetInsertMarkPosition, EM_SETINSERTMARK
	LRESULT OnSetInsertMark(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c EM_SETINSERTMARKCOLOR handler</em>
	///
	/// Will be called if the color of the control's insertion mark shall be changed.
	/// We use this handler for insertion mark support.
	///
	/// \sa OnGetInsertMarkColor, insertMark, OnSetInsertMark, EM_SETINSERTMARKCOLOR
	LRESULT OnSetInsertMarkColor(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c EM_SETTABSTOPS handler</em>
	///
	/// Will be called if the control's tab stops shall be changed.
	/// We use this handler to synchronize our settings.
	///
	/// \sa get_TabStops, get_TabWidth,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672104.aspx">EM_SETTABSTOPS</a>
	LRESULT OnSetTabStops(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c EN_ALIGN_*_EC handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user has changed the control's
	/// writing direction.
	/// We use this handler to raise the \c WritingDirectionChanged event.
	///
	/// \sa get_RightToLeft, Raise_WritingDirectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672108.aspx">EN_ALIGN_LTR_EC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672109.aspx">EN_ALIGN_RTL_EC</a>
	LRESULT OnReflectedAlign(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_CHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's content has changed.
	/// We use this handler to raise the \c TextChanged event.
	///
	/// \sa OnSetText, get_Text, Raise_TextChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672110.aspx">EN_CHANGE</a>
	LRESULT OnReflectedChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_ERRSPACE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control couldn't allocate enough
	/// memory to meet a specific request.
	/// We use this handler to raise the \c OutOfMemory event.
	///
	/// \sa Raise_OutOfMemory,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672112.aspx">EN_ERRSPACE</a>
	LRESULT OnReflectedErrSpace(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_MAXTEXT handler</em>
	///
	/// Will be called if the control's parent window is notified, that the text, that was entered into the
	/// control, got truncated.
	/// We use this handler to raise the \c TruncatedText event.
	///
	/// \sa Raise_TruncatedText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672115.aspx">EN_MAXTEXT</a>
	LRESULT OnReflectedMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_*SCROLL handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control is being scrolled.
	/// We use this handler to raise the \c Scrolling event.
	///
	/// \sa OnScroll, Raise_Scrolling,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672113.aspx">EN_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672118.aspx">EN_VSCROLL</a>
	LRESULT OnReflectedScroll(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_SETFOCUS handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has gained the keyboard
	/// focus.
	/// We use this handler to initialize IME.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672116.aspx">EN_SETFOCUS</a>
	LRESULT OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_UPDATE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's content is about to be
	/// drawn.
	/// We use this handler to raise the \c BeforeDrawText event.
	///
	/// \sa Raise_BeforeDrawText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672117.aspx">EN_UPDATE</a>
	LRESULT OnReflectedUpdate(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_AbortedDrag, EditCtlsLibU::_ITextBoxEvents::AbortedDrag, Raise_Drop,
	///       EndDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_AbortedDrag, EditCtlsLibA::_ITextBoxEvents::AbortedDrag, Raise_Drop,
	///       EndDrag
	/// \endif
	inline HRESULT Raise_AbortedDrag(void);
	/// \brief <em>Raises the \c BeforeDrawText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_BeforeDrawText, EditCtlsLibU::_ITextBoxEvents::BeforeDrawText
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_BeforeDrawText, EditCtlsLibA::_ITextBoxEvents::BeforeDrawText
	/// \endif
	inline HRESULT Raise_BeforeDrawText(void);
	/// \brief <em>Raises the \c BeginDrag event</em>
	///
	/// \param[in] firstChar The zero-based index of the first character of the text that the user wants to
	///            drag.
	/// \param[in] lastChar The zero-based index of the last character of the text that the user wants to
	///            drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_BeginDrag, EditCtlsLibU::_ITextBoxEvents::BeginDrag, BeginDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginRDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_BeginDrag, EditCtlsLibA::_ITextBoxEvents::BeginDrag, BeginDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginRDrag
	/// \endif
	inline HRESULT Raise_BeginDrag(LONG firstChar, LONG lastChar, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c BeginRDrag event</em>
	///
	/// \param[in] firstChar The zero-based index of the first character of the text that the user wants to
	///            drag.
	/// \param[in] lastChar The zero-based index of the last character of the text that the user wants to
	///            drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_BeginRDrag, EditCtlsLibU::_ITextBoxEvents::BeginRDrag, BeginDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_BeginRDrag, EditCtlsLibA::_ITextBoxEvents::BeginRDrag, BeginDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginDrag
	/// \endif
	inline HRESULT Raise_BeginRDrag(LONG firstChar, LONG lastChar, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_Click, EditCtlsLibU::_ITextBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_Click, EditCtlsLibA::_ITextBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the control from
	///                showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_ContextMenu, EditCtlsLibU::_ITextBoxEvents::ContextMenu,
	///       Raise_RClick
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_ContextMenu, EditCtlsLibA::_ITextBoxEvents::ContextMenu,
	///       Raise_RClick
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_DblClick, EditCtlsLibU::_ITextBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_DblClick, EditCtlsLibA::_ITextBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_DestroyedControlWindow,
	///       EditCtlsLibU::_ITextBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_DestroyedControlWindow,
	///       EditCtlsLibA::_ITextBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c DragMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_DragMouseMove, EditCtlsLibU::_ITextBoxEvents::DragMouseMove,
	///       Raise_MouseMove, Raise_OLEDragMouseMove, BeginDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_DragMouseMove, EditCtlsLibA::_ITextBoxEvents::DragMouseMove,
	///       Raise_MouseMove, Raise_OLEDragMouseMove, BeginDrag
	/// \endif
	inline HRESULT Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c Drop event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_Drop, EditCtlsLibU::_ITextBoxEvents::Drop, Raise_AbortedDrag,
	///       EndDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_Drop, EditCtlsLibA::_ITextBoxEvents::Drop, Raise_AbortedDrag,
	///       EndDrag
	/// \endif
	inline HRESULT Raise_Drop(void);
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
	///   \sa Proxy_ITextBoxEvents::Fire_KeyDown, EditCtlsLibU::_ITextBoxEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_KeyDown, EditCtlsLibA::_ITextBoxEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
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
	///   \sa Proxy_ITextBoxEvents::Fire_KeyPress, EditCtlsLibU::_ITextBoxEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_KeyPress, EditCtlsLibA::_ITextBoxEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp
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
	///   \sa Proxy_ITextBoxEvents::Fire_KeyUp, EditCtlsLibU::_ITextBoxEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_KeyUp, EditCtlsLibA::_ITextBoxEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MClick, EditCtlsLibU::_ITextBoxEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MClick, EditCtlsLibA::_ITextBoxEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MDblClick, EditCtlsLibU::_ITextBoxEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MDblClick, EditCtlsLibA::_ITextBoxEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MouseDown, EditCtlsLibU::_ITextBoxEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MouseDown, EditCtlsLibA::_ITextBoxEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MouseEnter, EditCtlsLibU::_ITextBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MouseEnter, EditCtlsLibA::_ITextBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MouseHover, EditCtlsLibU::_ITextBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MouseHover, EditCtlsLibA::_ITextBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MouseLeave, EditCtlsLibU::_ITextBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MouseLeave, EditCtlsLibA::_ITextBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MouseMove, EditCtlsLibU::_ITextBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MouseMove, EditCtlsLibA::_ITextBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MouseUp, EditCtlsLibU::_ITextBoxEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MouseUp, EditCtlsLibA::_ITextBoxEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_MouseWheel, EditCtlsLibU::_ITextBoxEvents::MouseWheel,
	///       Raise_MouseMove, EditCtlsLibU::ExtendedMouseButtonConstants, EditCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_MouseWheel, EditCtlsLibA::_ITextBoxEvents::MouseWheel,
	///       Raise_MouseMove, EditCtlsLibA::ExtendedMouseButtonConstants, EditCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLECompleteDrag, EditCtlsLibU::_ITextBoxEvents::OLECompleteDrag,
	///       Raise_OLEStartDrag, EditCtlsLibU::IOLEDataObject::GetData, OLEDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLECompleteDrag, EditCtlsLibA::_ITextBoxEvents::OLECompleteDrag,
	///       Raise_OLEStartDrag, EditCtlsLibA::IOLEDataObject::GetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect);
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
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragDrop, EditCtlsLibU::_ITextBoxEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragDrop, EditCtlsLibA::_ITextBoxEvents::OLEDragDrop,
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
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragEnter, EditCtlsLibU::_ITextBoxEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragEnter, EditCtlsLibA::_ITextBoxEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The potential drop target window's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragEnterPotentialTarget,
	///       EditCtlsLibU::_ITextBoxEvents::OLEDragEnterPotentialTarget, Raise_OLEDragLeavePotentialTarget,
	///       OLEDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragEnterPotentialTarget,
	///       EditCtlsLibA::_ITextBoxEvents::OLEDragEnterPotentialTarget, Raise_OLEDragLeavePotentialTarget,
	///       OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragLeave, EditCtlsLibU::_ITextBoxEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragLeave, EditCtlsLibA::_ITextBoxEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragLeavePotentialTarget,
	///       EditCtlsLibU::_ITextBoxEvents::OLEDragLeavePotentialTarget, Raise_OLEDragEnterPotentialTarget,
	///       OLEDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragLeavePotentialTarget,
	///       EditCtlsLibA::_ITextBoxEvents::OLEDragLeavePotentialTarget, Raise_OLEDragEnterPotentialTarget,
	///       OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragLeavePotentialTarget(void);
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
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragMouseMove, EditCtlsLibU::_ITextBoxEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEDragMouseMove, EditCtlsLibA::_ITextBoxEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c DROPEFFECT enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLEGiveFeedback, EditCtlsLibU::_ITextBoxEvents::OLEGiveFeedback,
	///       Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEGiveFeedback, EditCtlsLibA::_ITextBoxEvents::OLEGiveFeedback,
	///       Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors);
	/// \brief <em>Raises the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c TRUE, the user has pressed the \c ESC key since the last time
	///            this event was raised; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue (\c S_OK), cancel
	///                (\c DRAGDROP_S_CANCEL) or complete (\c DRAGDROP_S_DROP) the drag'n'drop operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLEQueryContinueDrag,
	///       EditCtlsLibU::_ITextBoxEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEQueryContinueDrag,
	///       EditCtlsLibA::_ITextBoxEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \endif
	inline HRESULT Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith);
	/// \brief <em>Raises the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLEReceivedNewData,
	///       EditCtlsLibU::_ITextBoxEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEReceivedNewData,
	///       EditCtlsLibA::_ITextBoxEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLESetData, EditCtlsLibU::_ITextBoxEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLESetData, EditCtlsLibA::_ITextBoxEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OLEStartDrag, EditCtlsLibU::_ITextBoxEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OLEStartDrag, EditCtlsLibA::_ITextBoxEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEStartDrag(IOLEDataObject* pData);
	/// \brief <em>Raises the \c OutOfMemory event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_OutOfMemory, EditCtlsLibU::_ITextBoxEvents::OutOfMemory
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_OutOfMemory, EditCtlsLibA::_ITextBoxEvents::OutOfMemory
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_RClick, EditCtlsLibU::_ITextBoxEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_RClick, EditCtlsLibA::_ITextBoxEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_RDblClick, EditCtlsLibU::_ITextBoxEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_RDblClick, EditCtlsLibA::_ITextBoxEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_RecreatedControlWindow,
	///       EditCtlsLibU::_ITextBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_RecreatedControlWindow,
	///       EditCtlsLibA::_ITextBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_ResizedControlWindow,
	///       EditCtlsLibU::_ITextBoxEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_ResizedControlWindow,
	///       EditCtlsLibA::_ITextBoxEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c Scrolling event</em>
	///
	/// \param[in] axis The axis which is scrolled. Any of the values defined by the \c ScrollAxisConstants
	///            enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event isn't raised if the \c MultiLine property is set to \c VARIANT_FALSE.\n
	///          This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_Scrolling, EditCtlsLibU::_ITextBoxEvents::Scrolling,
	///       EditCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_Scrolling, EditCtlsLibA::_ITextBoxEvents::Scrolling,
	///       EditCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_Scrolling(ScrollAxisConstants axis);
	/// \brief <em>Raises the \c TextChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default event.\n
	///          This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_TextChanged, EditCtlsLibU::_ITextBoxEvents::TextChanged
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_TextChanged, EditCtlsLibA::_ITextBoxEvents::TextChanged
	/// \endif
	inline HRESULT Raise_TextChanged(void);
	/// \brief <em>Raises the \c TruncatedText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_TruncatedText, EditCtlsLibU::_ITextBoxEvents::TruncatedText
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_TruncatedText, EditCtlsLibA::_ITextBoxEvents::TruncatedText
	/// \endif
	inline HRESULT Raise_TruncatedText(void);
	/// \brief <em>Raises the \c WritingDirectionChanged event</em>
	///
	/// \param[in] newWritingDirection The control's new writing direction. Any of the values defined by the
	///            \c WritingDirectionConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to limitations of Microsoft Windows, this event is not raised if the writing direction
	///          is changed using the control's default context menu.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_WritingDirectionChanged,
	///       EditCtlsLibU::_ITextBoxEvents::WritingDirectionChanged, EditCtlsLibU::WritingDirectionConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_WritingDirectionChanged,
	///       EditCtlsLibA::_ITextBoxEvents::WritingDirectionChanged, EditCtlsLibA::WritingDirectionConstants
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_XClick, EditCtlsLibU::_ITextBoxEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick,
	///       EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_XClick, EditCtlsLibA::_ITextBoxEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick,
	///       EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ITextBoxEvents::Fire_XDblClick, EditCtlsLibU::_ITextBoxEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick,
	///       EditCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ITextBoxEvents::Fire_XDblClick, EditCtlsLibA::_ITextBoxEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick,
	///       EditCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
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
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
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
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* and \c ES_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control by sending messages</em>
	///
	/// Sends \c WM_* and \c EM_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
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
	/// \brief <em>Auto-scrolls the control</em>
	///
	/// \sa OnTimer, Raise_DragMouseMove, DragDropStatus::AutoScrolling
	void AutoScroll(void);
	/// \brief <em>Retrieves whether the specified point is located over selected text</em>
	///
	/// \param[in] pt The point to check.
	///
	/// \return \c TRUE if the specified point is located over selected text; otherwise \c FALSE.
	///
	/// \sa get_SelectedText
	BOOL IsOverSelectedText(LPARAM pt);

	/// \brief <em>Draws the insertion mark</em>
	///
	/// This method is called by \c OnPaint to draw the control's insertion mark.
	///
	/// \param[in] targetDC The device context to be used for drawing.
	///
	/// \sa OnPaint
	void DrawInsertionMark(CDCHandle targetDC);
	/// \brief <em>Retrieves whether the logical left mouse button is held down</em>
	///
	/// \return \c TRUE if the logical left mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsRightMouseButtonDown
	BOOL IsLeftMouseButtonDown(void);
	/// \brief <em>Retrieves whether the logical right mouse button is held down</em>
	///
	/// \return \c TRUE if the logical right mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsLeftMouseButtonDown
	BOOL IsRightMouseButtonDown(void);


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
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
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
			CComObject< PropertyNotifySinkImpl<TextBox> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(TextBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<TextBox> >::CreateInstance(&pPropertyNotifySink);
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
			CComObject< PropertyNotifySinkImpl<TextBox> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(TextBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<TextBox> >::CreateInstance(&pPropertyNotifySink);
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

		/// \brief <em>The autoscroll-zone's default width</em>
		///
		/// The default width (in pixels) of the border around the control's client area, that's
		/// sensitive for auto-scrolling during a drag'n'drop operation. If the mouse cursor's position
		/// lies within this area during a drag'n'drop operation, the control will be auto-scrolled.
		///
		/// \sa dragScrollTimeBase, Raise_DragMouseMove
		static const int DRAGSCROLLZONEWIDTH = 16;

		/// \brief <em>Holds the \c AcceptNumbersOnly property's setting</em>
		///
		/// \sa get_AcceptNumbersOnly, put_AcceptNumbersOnly
		UINT acceptNumbersOnly : 1;
		/// \brief <em>Holds the \c AcceptTabKey property's setting</em>
		///
		/// \sa get_AcceptTabKey, put_AcceptTabKey
		UINT acceptTabKey : 1;
		/// \brief <em>Holds the \c AllowDragDrop property's setting</em>
		///
		/// \sa get_AllowDragDrop, put_AllowDragDrop
		UINT allowDragDrop : 1;
		/// \brief <em>Holds the \c AlwaysShowSelection property's setting</em>
		///
		/// \sa get_AlwaysShowSelection, put_AlwaysShowSelection
		UINT alwaysShowSelection : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c AutoScrolling property's setting</em>
		///
		/// \sa get_AutoScrolling, put_AutoScrolling
		AutoScrollingConstants autoScrolling;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
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
		/// \brief <em>Holds the \c DragScrollTimeBase property's setting</em>
		///
		/// \sa get_DragScrollTimeBase, put_DragScrollTimeBase
		long dragScrollTimeBase;
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
		/// \brief <em>Holds the \c FormattingRectangle* properties' settings</em>
		///
		/// \sa get_FormattingRectangleHeight, put_FormattingRectangleHeight,
		///     get_FormattingRectangleLeft, put_FormattingRectangleLeft,
		///     get_FormattingRectangleTop, put_FormattingRectangleTop,
		///     get_FormattingRectangleWidth, put_FormattingRectangleWidth
		CRect formattingRectangle;
		/// \brief <em>Holds the \c HAlignment property's setting</em>
		///
		/// \sa get_HAlignment, put_HAlignment
		HAlignmentConstants hAlignment;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c IMEMode property's setting</em>
		///
		/// \sa get_IMEMode, put_IMEMode
		IMEModeConstants IMEMode;
		/// \brief <em>Holds the \c InsertMarkColor property's setting</em>
		///
		/// \sa get_InsertMarkColor, put_InsertMarkColor
		OLE_COLOR insertMarkColor;
		/// \brief <em>Holds the \c InsertSoftLineBreaks property's setting</em>
		///
		/// \sa get_InsertSoftLineBreaks, put_InsertSoftLineBreaks
		UINT insertSoftLineBreaks : 1;
		/// \brief <em>Holds the \c LeftMargin property's setting</em>
		///
		/// \sa get_LeftMargin, put_LeftMargin
		OLE_XSIZE_PIXELS leftMargin;
		/// \brief <em>Holds the \c MaxTextLength property's setting</em>
		///
		/// \sa get_MaxTextLength, put_MaxTextLength
		long maxTextLength;
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
		/// \brief <em>Holds the \c MultiLine property's setting</em>
		///
		/// \sa get_MultiLine, put_MultiLine
		UINT multiLine : 1;
		/// \brief <em>Holds the \c OLEDragImageStyle property's setting</em>
		///
		/// \sa get_OLEDragImageStyle, put_OLEDragImageStyle
		OLEDragImageStyleConstants oleDragImageStyle;
		/// \brief <em>Holds the \c PasswordChar property's setting</em>
		///
		/// \sa get_PasswordChar, put_PasswordChar
		TCHAR passwordChar;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c ReadOnly property's setting</em>
		///
		/// \sa get_ReadOnly, put_ReadOnly
		UINT readOnly : 1;
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
		/// \brief <em>Holds the \c ScrollBars property's setting</em>
		///
		/// \sa get_ScrollBars, put_ScrollBars
		ScrollBarsConstants scrollBars;
		/// \brief <em>Holds the \c SelectedTextMouseIcon property's settings</em>
		///
		/// \sa get_SelectedTextMouseIcon, put_SelectedTextMouseIcon, putref_SelectedTextMouseIcon
		PictureProperty selectedTextMouseIcon;
		/// \brief <em>Holds the \c SelectedTextMousePointer property's setting</em>
		///
		/// \sa get_SelectedTextMousePointer, put_SelectedTextMousePointer
		MousePointerConstants selectedTextMousePointer;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		#ifdef USE_STL
			/// \brief <em>Holds the \c TabStops property's setting</em>
			///
			/// \sa get_TabStops, put_TabStops
			std::vector<long> tabStops;
		#else
			/// \brief <em>Holds the \c TabStops property's setting</em>
			///
			/// \sa get_TabStops, put_TabStops
			CAtlArray<long> tabStops;
		#endif
		/// \brief <em>Holds the \c TabWidth property's setting in dialog template units</em>
		///
		/// \sa get_TabWidth, put_TabWidth
		long tabWidthInDTUs;
		/// \brief <em>Holds the \c TabWidth property's setting in pixels</em>
		///
		/// \sa get_TabWidth, put_TabWidth
		long tabWidthInPixels;
		/// \brief <em>If \c TRUE, \c Create will set \c text to the control's name</em>
		///
		/// \sa Create, text
		UINT resetTextToName : 1;
		/// \brief <em>Holds the \c Text property's setting</em>
		///
		/// \sa get_Text, put_Text
		CComBSTR text;
		/// \brief <em>Holds the \c UseCustomFormattingRectangle property's setting</em>
		///
		/// \sa get_UseCustomFormattingRectangle, put_UseCustomFormattingRectangle
		UINT useCustomFormattingRectangle : 1;
		/// \brief <em>Holds the \c UsePasswordChar property's setting</em>
		///
		/// \sa get_UsePasswordChar, put_UsePasswordChar
		UINT usePasswordChar : 1;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;

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
			acceptTabKey = FALSE;
			allowDragDrop = TRUE;
			alwaysShowSelection = FALSE;
			appearance = a3D;
			autoScrolling = asHorizontal;
			backColor = 0x80000000 | COLOR_WINDOW;
			borderStyle = bsNone;
			cancelIMECompositionOnSetFocus = FALSE;
			characterConversion = ccNone;
			completeIMECompositionOnKillFocus = FALSE;
			cueBanner = L"";
			disabledBackColor = static_cast<OLE_COLOR>(-1);
			disabledEvents = static_cast<DisabledEventsConstants>(deBeforeDrawText | deClickEvents | deMouseEvents | deScrolling);
			disabledForeColor = static_cast<OLE_COLOR>(-1);
			displayCueBannerOnFocus = FALSE;
			dontRedraw = FALSE;
			doOEMConversion = FALSE;
			dragScrollTimeBase = -1;
			enabled = TRUE;
			foreColor = 0x80000000 | COLOR_WINDOWTEXT;
			ZeroMemory(&formattingRectangle, sizeof(formattingRectangle));
			hAlignment = halLeft;
			hoverTime = -1;
			IMEMode = imeInherit;
			insertMarkColor = RGB(0, 0, 0);
			insertSoftLineBreaks = FALSE;
			leftMargin = -1;
			maxTextLength = -1;
			modified = FALSE;
			mousePointer = mpDefault;
			multiLine = FALSE;
			oleDragImageStyle = odistClassic;
			passwordChar = 0;
			processContextMenuKeys = TRUE;
			readOnly = FALSE;
			registerForOLEDragDrop = FALSE;
			rightMargin = -1;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			scrollBars = static_cast<ScrollBarsConstants>(0);
			selectedTextMousePointer = mpDefault;
			supportOLEDragImages = TRUE;
			#ifdef USE_STL
				tabStops.clear();
			#else
				tabStops.RemoveAll();
			#endif
			tabWidthInDTUs = -1;
			tabWidthInPixels = -1;
			text = L"TextBox";
			resetTextToName = TRUE;
			useCustomFormattingRectangle = FALSE;
			usePasswordChar = FALSE;
			useSystemFont = TRUE;
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
		/// \sa OnMouseActivate, OnSetFocus, OnKillFocus
		UINT uiActivationPending : 1;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;
		/// \brief <em>If \c TRUE, \c AppendText may invalidate the window and \c OnReflectedCtlColorStatic will set the background mode to \c TRANSPARENT</em>
		///
		/// If the text background mode is set to \c TRANSPARENT in \c OnReflectedCtlColorStatic and
		/// \c AppendText scrolls the control, the control is not redrawn properly. But if the control's
		/// background isn't a solid color (e. g. when it sits on a themed tab strip), the text background
		/// <strong>must</strong> be set to \c TRANSPARENT. Therefore, we use this flag to a) set the correct
		/// background mode in \c OnReflectedCtlColorStatic and b) force a refresh of the control in
		/// \c AppendText, if necessary.
		///
		/// \sa AppendText, OnReflectedCtlColorStatic
		UINT useTransparentTextBackground : 1;

		Flags()
		{
			uiActivationPending = FALSE;
			usingThemes = FALSE;
			useTransparentTextBackground = FALSE;
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

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

	/// \brief <em>Holds the position of the control's insertion mark</em>
	struct InsertMark
	{
		/// \brief <em>If set to \c TRUE, the insertion mark won't be drawn</em>
		///
		/// \sa OnScroll, DrawInsertionMark
		UINT hidden : 1;
		/// \brief <em>The zero-based index of the character at which the insertion mark is placed</em>
		int characterIndex;
		/// \brief <em>The insertion mark's position relative to the character</em>
		BOOL afterChar;
		/// \brief <em>The insertion mark's color</em>
		COLORREF color;

		InsertMark()
		{
			Reset();
		}

		/// \brief <em>Resets all member variables</em>
		void Reset(void)
		{
			hidden = FALSE;
			characterIndex = -1;
			afterChar = FALSE;
			color = RGB(0, 0, 0);
		}

		/// \brief <em>Processes the parameters of a \c EM_SETINSERTMARK message to store the new insertion mark</em>
		///
		/// \sa OnSetInsertMark
		void ProcessSetInsertMark(WPARAM wParam, LPARAM lParam)
		{
			afterChar = static_cast<BOOL>(wParam);
			characterIndex = static_cast<int>(lParam);
		}
	} insertMark;

	//////////////////////////////////////////////////////////////////////
	/// \name Drag'n'Drop
	///
	//@{
	/// \brief <em>The \c CLSID_WICImagingFactory object used to create WIC objects that are required during drag image creation</em>
	///
	/// \sa OnGetDragImage, CreateThumbnail
	CComPtr<IWICImagingFactory> pWICImagingFactory;
	/// \brief <em>Creates a thumbnail of the specified icon in the specified size</em>
	///
	/// \param[in] hIcon The icon to create the thumbnail for.
	/// \param[in] size The thumbnail's size in pixels.
	/// \param[in,out] pBits The thumbnail's DIB bits.
	/// \param[in] doAlphaChannelPostProcessing WIC has problems to handle the alpha channel of the icon
	///            specified by \c hIcon. If this parameter is set to \c TRUE, some post-processing is done
	///            to correct the pixel failures. Otherwise the failures are not corrected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnGetDragImage, pWICImagingFactory
	HRESULT CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing);

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>The zero-based index of the first character of the dragged text</em>
		int draggedTextFirstChar;
		/// \brief <em>The zero-based index of the last character of the dragged text</em>
		int draggedTextLastChar;
		/// \brief <em>The handle of the imagelist containing the drag image</em>
		///
		/// \sa get_hDragImageList
		HIMAGELIST hDragImageList;
		/// \brief <em>Enables or disables auto-destruction of \c hDragImageList</em>
		///
		/// Controls whether the imagelist defined by \c hDragImageList is destroyed automatically. If set to
		/// \c TRUE, it is destroyed in \c EndDrag; otherwise not.
		///
		/// \sa hDragImageList, EndDrag
		UINT autoDestroyImgLst : 1;
		/// \brief <em>Indicates whether the drag image is visible or hidden</em>
		///
		/// If this value is 0, the drag image is visible; otherwise not.
		///
		/// \sa get_hDragImageList, get_ShowDragImage, put_ShowDragImage, ShowDragImage, HideDragImage,
		///     IsDragImageVisible
		int dragImageIsHidden;

		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>The currently dragged data for the case that the we're the drag source</em>
		CComPtr<IDataObject> pSourceDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the mouse button (as \c MK_* constant) that the drag'n'drop operation is performed with</em>
		DWORD draggingMouseButton;
		/// \brief <em>If \c TRUE, we'll hide and re-show the drag image in \c IDropTarget::DragEnter so that the item count label is displayed</em>
		///
		/// \sa DragEnter, OLEDrag
		UINT useItemCountLabelHack : 1;
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

		/// \brief <em>Holds data and flags related to auto-scrolling</em>
		///
		/// \sa AutoScroll
		struct AutoScrolling
		{
			/// \brief <em>Holds the current speed multiplier used for horizontal auto-scrolling</em>
			LONG currentHScrollVelocity;
			/// \brief <em>Holds the current speed multiplier used for vertical auto-scrolling</em>
			LONG currentVScrollVelocity;
			/// \brief <em>Holds the current interval of the auto-scroll timer</em>
			LONG currentTimerInterval;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled downwards</em>
			DWORD lastScroll_Down;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the left</em>
			DWORD lastScroll_Left;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the right</em>
			DWORD lastScroll_Right;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled upwardly</em>
			DWORD lastScroll_Up;

			AutoScrolling()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				currentHScrollVelocity = 0;
				currentVScrollVelocity = 0;
				currentTimerInterval = 0;
				lastScroll_Down = 0;
				lastScroll_Left = 0;
				lastScroll_Right = 0;
				lastScroll_Up = 0;
			}
		} autoScrolling;

		/// \brief <em>Holds data required for drag-detection</em>
		struct Candidate
		{
			/// \brief <em>The zero-based index of the first character of the text, that might be dragged</em>
			int textFirstChar;
			/// \brief <em>The zero-based index of the last character of the text, that might be dragged</em>
			int textLastChar;
			/// \brief <em>The position in pixels at which dragging might start</em>
			POINT position;
			/// \brief <em>The mouse button (as \c MK_* constant) that the drag'n'drop operation might be performed with</em>
			int button;

			/// \brief <em>Holds the parameters of a buffered \c WM_*BUTTONDOWN message</em>
			typedef struct BufferedMessage
			{
				/// \brief <em>Holds the \c wParam parameter of the buffered message</em>
				WPARAM wParam;
				/// \brief <em>Holds the \c lParam parameter of the buffered message</em>
				LPARAM lParam;

				BufferedMessage()
				{
					Reset();
				}

				/// \brief <em>Resets all member variables to their defaults</em>
				void Reset(void)
				{
					wParam = 0;
					lParam = 0;
				}
			} BufferedMessage;
			/// \brief <em>Holds the parameters of a buffered \c WM_LBUTTONDOWN message</em>
			BufferedMessage bufferedLButtonDown;
			/// \brief <em>Holds the parameters of a buffered \c WM_RBUTTONDOWN message</em>
			BufferedMessage bufferedRButtonDown;

			Candidate()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				button = 0;
				textFirstChar = -1;
				textLastChar = -1;
				bufferedLButtonDown.Reset();
				bufferedRButtonDown.Reset();
			}
		} candidate;

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pSourceDataObject = NULL;
			pDropTargetHelper = NULL;
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;

			draggedTextFirstChar = -1;
			draggedTextLastChar = -1;
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
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
			this->draggedTextFirstChar = -1;
			this->draggedTextLastChar = -1;
			if(this->hDragImageList && autoDestroyImgLst) {
				ImageList_Destroy(this->hDragImageList);
			}
			this->hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;

			candidate.Reset();

			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			if(this->pSourceDataObject) {
				this->pSourceDataObject = NULL;
			}
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Decrements the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, HideDragImage, IsDragImageVisible
		void ShowDragImage(BOOL commonDragDropOnly)
		{
			if(this->hDragImageList) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					ImageList_DragShowNolock(TRUE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					pDropTargetHelper->Show(TRUE);
				}
			}
		}

		/// \brief <em>Increments the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, ShowDragImage, IsDragImageVisible
		void HideDragImage(BOOL commonDragDropOnly)
		{
			if(this->hDragImageList) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					ImageList_DragShowNolock(FALSE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					pDropTargetHelper->Show(FALSE);
				}
			}
		}

		/// \brief <em>Retrieves whether we're currently displaying a drag image</em>
		///
		/// \return \c TRUE, if we're displaying a drag image; otherwise \c FALSE.
		///
		/// \sa dragImageIsHidden, ShowDragImage, HideDragImage
		BOOL IsDragImageVisible(void)
		{
			return (dragImageIsHidden == 0);
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation started</em>
		///
		/// \param[in] pTextBox The \c TextBox control whose \c CreateLegacyDragImage method shall be called.
		/// \param[in] hWndEdit The edit control window, that the method will work on to calculate the position
		///            of the drag image's hotspot.
		/// \param[in] draggedTextFirstChr The zero-based index of the first character of the text to drag.
		/// \param[in] draggedTextLastChr The zero-based index of the last character of the text to drag.
		/// \param[in] hDragImgLst The imagelist containing the drag image that shall be used to
		///            visualize the drag'n'drop operation. If -1, the method will create the drag image
		///            itself; if \c NULL, no drag image will be displayed.
		/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImgLst parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImgLst parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImgLst parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImgLst parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa EndDrag, TextBox::CreateLegacyDragImage
		HRESULT BeginDrag(TextBox* pTextBox, HWND hWndEdit, LONG draggedTextFirstChr, LONG draggedTextLastChr, HIMAGELIST hDragImgLst, int* pXHotSpot, int* pYHotSpot)
		{
			ATLASSUME(pTextBox);

			UINT b = FALSE;
			if(hDragImgLst == static_cast<HIMAGELIST>(LongToHandle(-1))) {
				POINT upperLeftPoint = {0};
				hDragImgLst = pTextBox->CreateLegacyDragImage(draggedTextFirstChr, draggedTextLastChr, &upperLeftPoint, NULL);
				if(!hDragImgLst) {
					return E_FAIL;
				}
				b = TRUE;

				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				::ScreenToClient(hWndEdit, &mousePosition);
				if(CWindow(hWndEdit).GetExStyle() & WS_EX_LAYOUTRTL) {
					SIZE dragImageSize = {0};
					ImageList_GetIconSize(hDragImgLst, reinterpret_cast<PINT>(&dragImageSize.cx), reinterpret_cast<PINT>(&dragImageSize.cy));
					*pXHotSpot = upperLeftPoint.x + dragImageSize.cx - mousePosition.x;
				} else {
					*pXHotSpot = mousePosition.x - upperLeftPoint.x;
				}
				*pYHotSpot = mousePosition.y - upperLeftPoint.y;
			}

			if(this->hDragImageList && this->autoDestroyImgLst) {
				ImageList_Destroy(this->hDragImageList);
			}

			this->autoDestroyImgLst = b;
			this->hDragImageList = hDragImgLst;
			this->draggedTextFirstChar = draggedTextFirstChr;
			this->draggedTextLastChar = draggedTextLastChr;

			dragImageIsHidden = 1;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation ended</em>
		///
		/// \sa BeginDrag
		void EndDrag(void)
		{
			this->draggedTextFirstChar = -1;
			this->draggedTextLastChar = -1;
			if(autoDestroyImgLst && this->hDragImageList) {
				ImageList_Destroy(this->hDragImageList);
			}
			this->hDragImageList = NULL;
			dragImageIsHidden = 1;
			autoScrolling.Reset();
		}

		/// \brief <em>Retrieves whether we're in drag'n'drop mode</em>
		///
		/// \return \c TRUE if we're in drag'n'drop mode; otherwise \c FALSE.
		///
		/// \sa BeginDrag, EndDrag
		BOOL IsDragging(void)
		{
			return (this->draggedTextFirstChar != -1 && this->draggedTextLastChar != -1);
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			autoScrolling.Reset();
		}
	} dragDropStatus;
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-scroll the control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;

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
	/// \sa OnChar, OnKeyDown, OnKeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	WPARAM cachedWParam;

private:
};     // TextBox

OBJECT_ENTRY_AUTO(__uuidof(TextBox), TextBox)