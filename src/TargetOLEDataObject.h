//////////////////////////////////////////////////////////////////////
/// \class TargetOLEDataObject
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communicates with an object through the \c IDataObject interface</em>
///
/// This class provides easy access to an object that implements the \c IDataObject interface. It is
/// used if the control is the target of an OLE drag'n'drop operation.
///
/// \if UNICODE
///   \sa EditCtlsLibU::IOLEDataObject, SourceOLEDataObject, TextBox, UpDownTextBox,
///       <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
/// \else
///   \sa EditCtlsLibA::IOLEDataObject, SourceOLEDataObject, TextBox, UpDownTextBox,
///       <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "EditCtlsU.h"
#else
	#include "EditCtlsA.h"
#endif
#include "_IOLEDataObjectEvents_CP.h"
#include "helpers.h"
#include "TextBox.h"
#include "UpDownTextBox.h"


class ATL_NO_VTABLE TargetOLEDataObject : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TargetOLEDataObject, &CLSID_OLEDataObject>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TargetOLEDataObject>,
    public Proxy_IOLEDataObjectEvents<TargetOLEDataObject>,
    #ifdef UNICODE
    	public IDispatchImpl<IOLEDataObject, &IID_IOLEDataObject, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IOLEDataObject, &IID_IOLEDataObject, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class TextBox;
	friend class UpDownTextBox;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_OLEDATAOBJECT)

		BEGIN_COM_MAP(TargetOLEDataObject)
			COM_INTERFACE_ENTRY(IOLEDataObject)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(IID_IDataObject, 0, QueryIDataObjectInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TargetOLEDataObject)
			CONNECTION_POINT_ENTRY(__uuidof(_IOLEDataObjectEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
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

	/// \brief <em>Will be called if this class is \c QueryInterface()'d for \c IDataObject</em>
	///
	/// This method will be called if the class's \c QueryInterface() method is called with
	/// \c IID_IDataObject. We forward the call to the wrapped \c IDataObject implementation.
	///
	/// \param[in] pThis The instance of this class, that the interface is queried from.
	/// \param[in] queriedInterface Should be \c IID_IDataObject.
	/// \param[out] ppImplementation Receives the wrapped object's \c IDataObject implementation.
	/// \param[in] cookie A \c DWORD value specified in the COM interface map.
	///
	/// \return An \c HRESULT error code.
	static HRESULT CALLBACK QueryIDataObjectInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOLEDataObject
	///
	//@{
	/// \brief <em>Deletes the data hold by the \c IDataObject implementation</em>
	///
	/// \return Raises VB error 425 (invalid object use).
	///
	/// \sa SetData
	virtual HRESULT STDMETHODCALLTYPE Clear(void);
	/// \brief <em>Retrieves the best format settings from the \c IDataObject implementation</em>
	///
	/// Call this method to retrieve data format settings that match best with the data format, that you
	/// actually want to work with. Set the parameters to the data format settings you want to work with.
	/// The method will set them to the settings you should use.
	///
	/// \param[in,out] pFormatID An integer value specifying the data format. Valid values are those defined
	///                by VB's \c ClipBoardConstants enumeration, but also any other format that was
	///                registered using the \c RegisterClipboardFormat API function.
	/// \param[in,out] pIndex An integer value that is assigned to the internal \c FORMATETC struct's
	///                \c lindex member. Usually you pass -1 here, but some formats like
	///                \c CFSTR_FILECONTENTS require multiple \c FORMATETC structs for the same format. In
	///                such cases you'll give each struct of this format a separate index.
	/// \param[in,out] pDataOrViewAspect An integer value that is assigned to the internal \c FORMATETC
	///                struct's \c dwAspect member. Any of the \c DVASPECT_* values defined by the
	///                Microsoft&reg; Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, GetFormat,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	virtual HRESULT STDMETHODCALLTYPE GetCanonicalFormat(LONG* pFormatID, LONG* pIndex, LONG* pDataOrViewAspect);
	/// \brief <em>Retrieves data from the \c IDataObject implementation</em>
	///
	/// Retrieves data from the \c IDataObject implementation, that has the specified format.
	///
	/// \param[in] formatID An integer value specifying the format of the data to retrieve. Valid values are
	///            those defined by VB's \c ClipBoardConstants enumeration, but also any other format that
	///            has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually you pass -1 here, but some formats like \c CFSTR_FILECONTENTS require
	///            multiple \c FORMATETC structs for the same format. In such cases you'll give each struct
	///            of this format a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid.
	/// \param[out] pData The data that has the specified format.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method will fail, if the \c IDataObject implementation does not contain data of the
	///          specified format.
	///
	/// \sa GetCanonicalFormat, GetFormat, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	virtual HRESULT STDMETHODCALLTYPE GetData(LONG formatID, LONG index = -1, LONG dataOrViewAspect = DVASPECT_CONTENT, VARIANT* pData = NULL);
	/// \brief <em>Retrieves the \c DROPDESCRIPTION data stored by the \c IDataObject implementation</em>
	///
	/// Retrieves the \c DROPDESCRIPTION data stored by the \c IDataObject implementation. The drop
	/// description describes what will happen if the user drops the dragged data at the current position. It
	/// is displayed at the bottom of the drag image.
	///
	/// \param[in,out] pTargetDescription Receives the description of the current drop target.
	/// \param[in,out] pActionDescription Receives the description of the whole drop action, i. e. a string
	///                like <em>"Copy to %1"</em> where <em>"Copy to"</em> is the description of the current
	///                drop effect and <em>"%1"</em> is the placeholder for the drop target description
	///                specified by \c pTargetDescription.
	/// \param[in,out] pIcon Receives the icon used to visualize the current drop effect. Any of the values
	///                defined by the \c DropDescriptionIconConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks In the \c pActionDescription string, the sign "%" is escaped as "%%".\n
	///          Requires Windows Vista or newer.
	///
	/// \if UNICODE
	///   \sa SetDropDescription, HotKeyBox::get_SupportOLEDragImages,
	///       IPAddressBox::get_SupportOLEDragImages, TextBox::get_SupportOLEDragImages,
	///       UpDownTextBox::get_SupportOLEDragImages, EditCtlsLibU::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \else
	///   \sa SetDropDescription, HotKeyBox::get_SupportOLEDragImages,
	///       IPAddressBox::get_SupportOLEDragImages, TextBox::get_SupportOLEDragImages,
	///       UpDownTextBox::get_SupportOLEDragImages, EditCtlsLibA::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetDropDescription(VARIANT* pTargetDescription = NULL, VARIANT* pActionDescription = NULL, DropDescriptionIconConstants* pIcon = NULL);
	/// \brief <em>Retrieves whether the \c IDataObject implementation holds data in a specific format</em>
	///
	/// \param[in] formatID An integer value specifying the format to check for. Valid values are those
	///            defined by VB's \c ClipBoardConstants enumeration, but also any other format that was
	///            registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually you pass -1 here, but some formats like \c CFSTR_FILECONTENTS require
	///            multiple \c FORMATETC structs for the same format. In such cases you'll give each struct
	///            of this format a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid.
	/// \param[out] pFormatAvailable If \c VARIANT_TRUE, the \c IDataObject implementation holds data in the
	///             specified format; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetCanonicalFormat, GetData, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	virtual HRESULT STDMETHODCALLTYPE GetFormat(LONG formatID, LONG index = -1, LONG dataOrViewAspect = DVASPECT_CONTENT, VARIANT_BOOL* pFormatAvailable = NULL);
	/// \brief <em>Transfers data to the \c IDataObject implementation</em>
	///
	/// \param[in] formatID An integer value specifying the format of the data being passed. Valid values are
	///            those defined by VB's \c ClipBoardConstants enumeration, but also any other format that
	///            has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] data A \c VARIANT value specifying the data to transfer. If not specified, the method will
	///            fail.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually you pass -1 here, but some formats like \c CFSTR_FILECONTENTS require
	///            multiple \c FORMATETC structs for the same format. In such cases you'll give each struct
	///            of this format a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid.
	///
	/// \sa GetData, Clear,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	virtual HRESULT STDMETHODCALLTYPE SetData(LONG formatID, VARIANT data = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), LONG index = -1, LONG dataOrViewAspect = DVASPECT_CONTENT);
	/// \brief <em>Sets the drop description displayed below the drag image</em>
	///
	/// Sets the \c DROPDESCRIPTION data. The drop description describes what will happen if the user drops
	/// the dragged data at the current position. It is displayed at the bottom of the drag image.
	///
	/// \param[in] targetDescription The description of the current drop target.
	/// \param[in] actionDescription The description of the whole drop action, i. e. a string like <em>"Copy
	///            to %1"</em> where <em>"Copy to"</em> is the description of the current drop effect and
	///            <em>"%1"</em> is the placeholder for the drop target description specified by
	///            \c targetDescription.
	/// \param[in] icon The icon used to visualize the current drop effect. Any of the values defined by the
	///            \c DropDescriptionIconConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks In the \c actionDescription string, the sign "%" is escaped as "%%".\n
	///          If no parameters are specified, the drop description is removed from the drag image.\n
	///          Requires Windows Vista or newer.
	///
	/// \if UNICODE
	///   \sa GetDropDescription, HotKeyBox::put_SupportOLEDragImages,
	///       IPAddressBox::put_SupportOLEDragImages, TextBox::put_SupportOLEDragImages,
	///       UpDownTextBox::put_SupportOLEDragImages, EditCtlsLibU::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \else
	///   \sa GetDropDescription, HotKeyBox::put_SupportOLEDragImages,
	///       IPAddressBox::put_SupportOLEDragImages, TextBox::put_SupportOLEDragImages,
	///       UpDownTextBox::put_SupportOLEDragImages, EditCtlsLibA::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetDropDescription(VARIANT targetDescription = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), VARIANT actionDescription = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), DropDescriptionIconConstants icon = ddiNone);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given \c IDataObject implementation</em>
	///
	/// Attaches this object to a given \c IDataObject implementation, so that the implementation's methods
	/// can be called using this object's methods.
	///
	/// \param[in] pDataObject The \c IDataObject implementation to attach to.
	///
	/// \sa Detach
	void Attach(LPDATAOBJECT pDataObject);
	/// \brief <em>Detaches this object from an \c IDataObject implementation</em>
	///
	/// Detaches this object from the \c IDataObject implementation it currently wraps, so that it doesn't
	/// wrap any \c IDataObject implementation anymore.
	///
	/// \sa Attach
	void Detach(void);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c IDataObject implementation of the object we're working on</em>
		CComPtr<IDataObject> pDataObject;

		Properties()
		{
			pDataObject = NULL;
		}
	} properties;
};     // TargetOLEDataObject

OBJECT_ENTRY_AUTO(__uuidof(OLEDataObject), TargetOLEDataObject)