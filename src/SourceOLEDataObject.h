//////////////////////////////////////////////////////////////////////
/// \class SourceOLEDataObject
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communicates with an object through the \c IDataObject interface</em>
///
/// This class provides easy access to an object that implements the \c IDataObject interface. It is
/// used if the control is the source of an OLE drag'n'drop operation.
///
/// \if UNICODE
///   \sa EditCtlsLibU::IOLEDataObject, TargetOLEDataObject, TextBox,
///       <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
/// \else
///   \sa EditCtlsLibA::IOLEDataObject, TargetOLEDataObject, TextBox,
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


class ATL_NO_VTABLE SourceOLEDataObject : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<SourceOLEDataObject, &CLSID_OLEDataObject>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<SourceOLEDataObject>,
    public Proxy_IOLEDataObjectEvents<SourceOLEDataObject>,
    #ifdef UNICODE
    	public IDispatchImpl<IOLEDataObject, &IID_IOLEDataObject, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IOLEDataObject, &IID_IOLEDataObject, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IDataObject,
    public IEnumFORMATETC
{
	friend class TextBox;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_OLEDATAOBJECT)

		BEGIN_COM_MAP(SourceOLEDataObject)
			COM_INTERFACE_ENTRY(IOLEDataObject)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IDataObject)
			COM_INTERFACE_ENTRY(IEnumFORMATETC)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(SourceOLEDataObject)
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

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOLEDataObject
	///
	//@{
	/// \brief <em>Deletes the data hold by the \c SourceOLEDataObject object</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetData
	virtual HRESULT STDMETHODCALLTYPE Clear(void);
	/// \brief <em>Retrieves the best format settings from the \c SourceOLEDataObject object</em>
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
	/// \brief <em>Retrieves data from the \c SourceOLEDataObject object</em>
	///
	/// Retrieves data from the \c SourceOLEDataObject object, that has the specified format.
	///
	/// \param[in] formatID An integer value specifying the format of the data to retrieve. Valid values
	///            are those defined by VB's \c ClipBoardConstants enumeration, but also any other format
	///            that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c lindex member. Usually you pass -1 here, but some formats like
	///            \c CFSTR_FILECONTENTS require multiple \c FORMATETC structs for the same format. In
	///            such cases you'll give each struct of this format a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid.
	/// \param[out] pData The data that has the specified format.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method will fail, if the \c SourceOLEDataObject object does not contain data of
	///          the specified format.
	///
	/// \sa GetCanonicalFormat, GetFormat, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	virtual HRESULT STDMETHODCALLTYPE GetData(LONG formatID, LONG index = -1, LONG dataOrViewAspect = DVASPECT_CONTENT, VARIANT* pData = NULL);
	/// \brief <em>Retrieves the \c DROPDESCRIPTION data stored by the \c SourceOLEDataObject object</em>
	///
	/// Retrieves the \c DROPDESCRIPTION data stored by the \c SourceOLEDataObject object. The drop
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
	///   \sa SetDropDescription, TextBox::get_SupportOLEDragImages,
	///       EditCtlsLibU::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \else
	///   \sa SetDropDescription, TextBox::get_SupportOLEDragImages,
	///       EditCtlsLibA::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetDropDescription(VARIANT* pTargetDescription = NULL, VARIANT* pActionDescription = NULL, DropDescriptionIconConstants* pIcon = NULL);
	/// \brief <em>Retrieves whether the \c SourceOLEDataObject object holds data in a specific format</em>
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
	/// \param[out] pFormatAvailable If \c VARIANT_TRUE, the \c SourceOLEDataObject object holds data in the
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
	/// \brief <em>Transfers data to the \c SourceOLEDataObject object</em>
	///
	/// \param[in] formatID An integer value specifying the format of the data being passed. Valid values are
	///            those defined by VB's \c ClipBoardConstants enumeration, but also any other format that
	///            has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] data A \c VARIANT value specifying the data to transfer. If not specified, the
	///            \c OLESetData event will be raised if data of the specified format is requested from the
	///            \c SourceOLEDataObject object.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually you pass -1 here, but some formats like \c CFSTR_FILECONTENTS require
	///            multiple \c FORMATETC structs for the same format. In such cases you'll give each struct
	///            of this format a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid.
	///
	/// \sa GetData, Clear, TextBox::Raise_OLESetData,
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
	///   \sa GetDropDescription, TextBox::put_SupportOLEDragImages,
	///       EditCtlsLibU::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \else
	///   \sa GetDropDescription, TextBox::put_SupportOLEDragImages,
	///       EditCtlsLibA::DropDescriptionIconConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb773268.aspx">DROPDESCRIPTION</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetDropDescription(VARIANT targetDescription = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), VARIANT actionDescription = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), DropDescriptionIconConstants icon = ddiNone);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this object</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerTxtBox
	void SetOwner(__in_opt TextBox* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDataObject
	///
	//@{
	/// \brief <em>An object wants to get notified on changes to the data object</em>
	///
	/// This method is called by objects that want to track changes to the data contained by this object.
	/// The implementation should establish an advisory connection through the \c IAdviseSink interface.
	///
	/// \param[in] pFormatToTrack A \c FORMATETC struct specifying the data format that the object wants to
	///            track.
	/// \param[in] flags A bit field of flags (defined by the \c ADVF enumeration) controlling the advisory
	///            connection.
	/// \param[in] pAdviseSink The \c IAdviseSink implementation that the notifications should be sent to.
	/// \param[out] pConnectionID An integer value used to identify the established connection in
	///            \c DUnadvise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DUnadvise, EnumDAdvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692579.aspx">IDataObject::DAdvise</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693742.aspx">ADVF</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692513.aspx">IAdviseSink</a>
	virtual HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC* pFormatToTrack, DWORD flags, IAdviseSink* pAdviseSink, LPDWORD pConnectionID);
	/// \brief <em>An object no longer wants to receive notifications about changes to the data object</em>
	///
	/// This method is called by objects that called \c DAdvise to get notified of changes to the data
	/// contained by this object and now want to deregister. The implementation will close the connection.
	///
	/// \param[in] connectionID An integer value used to identify the connection to close. This value was
	///            returned previously by the \c DAdvise method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DAdvise, EnumDAdvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692448.aspx">IDataObject::DUnadvise</a>
	virtual HRESULT STDMETHODCALLTYPE DUnadvise(DWORD connectionID);
	/// \brief <em>Retrieves an enumerator to enumerate advisory connections</em>
	///
	/// Retrieves a \c STATDATA enumerator that may be used to iterate the currently open advisory connection
	/// objects.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumSTATDATA implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DAdvise, DUnadvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680127.aspx">IDataObject::EnumDAdvise</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688589.aspx">IEnumSTATDATA</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694341.aspx">STATDATA</a>
	virtual HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA** ppEnumerator);
	/// \brief <em>Retrieves an enumerator to enumerate the formats of the data contained by this object</em>
	///
	/// Retrieves a \c FORMATETC enumerator that may be used to iterate the data formats of the data
	/// contained by this object.
	///
	/// \param[in] direction The directions of data transfer to include in the enumeration. Any combination
	///            of the values defined by the \c DATADIR enumeration is valid.
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumFORMATETC implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, QueryGetData, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683979.aspx">IDataObject::EnumFormatEtc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682337.aspx">IEnumFORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680661.aspx">DATADIR</a>
	virtual HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD direction, IEnumFORMATETC** ppEnumerator);
	/// \brief <em>Translates a \c FORMATETC structure into a more general one</em>
	///
	/// Translates a \c FORMATETC structure into a more general one.
	///
	/// \param[in] pSpecifiedFormat The \c FORMATETC structure, that defines the format, medium and target
	///            device that the caller of this method would like to use in a subsequent call such as
	///            \c GetData.
	/// \param[out] pGeneralFormat The implementer's most general \c FORMATETC structure, that is canonically
	///             equivalent to \c pSpecifiedFormat.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, EnumFormatEtc, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680685.aspx">IDataObject::GetCanonicalFormatEtc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC* pSpecifiedFormat, FORMATETC* pGeneralFormat);
	/// \brief <em>Retrieves data from the data object</em>
	///
	/// Retrieves data of the specified format from this object.
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format, medium and target device to
	///            use when passing the data. It is possible to specify more than one medium by using the
	///            Boolean OR operator, allowing the method to choose the best medium among those specified.
	/// \param[out] pData The retrieved data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDataHere, QueryGetData, EnumFormatEtc, QueryGetData, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms678431.aspx">IDataObject::GetData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683812.aspx">STGMEDIUM</a>
	virtual HRESULT STDMETHODCALLTYPE GetData(FORMATETC* pFormat, STGMEDIUM* pData);
	/// \brief <em>Retrieves data from the data object</em>
	///
	/// Retrieves data of the specified format from this object. This method differs from \c GetData in that
	/// the caller must allocate and free the specified storage medium.
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format, medium and target device to
	///            use when passing the data. It is possible to specify more than one medium by using the
	///            Boolean OR operator, allowing the method to choose the best medium among those specified.
	/// \param[out] pData The retrieved data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, QueryGetData, EnumFormatEtc, QueryGetData, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687266.aspx">IDataObject::GetDataHere</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683812.aspx">STGMEDIUM</a>
	virtual HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC* /*pFormat*/, STGMEDIUM* /*pData*/);
	/// \brief <em>Determines whether this object contains data in a specific format</em>
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format, medium and target device to
	///            use for the query.
	///
	/// \return An \c HRESULT error code. \c S_OK, if the data object contains data in the format specified
	///         by the \c pFormat parameter.
	///
	/// \sa GetData, GetDataHere, EnumFormatEtc, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680637.aspx">IDataObject::QueryGetData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* pFormat);
	/// \brief <em>Transfers data to the data object</em>
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format of the data passed in the
	///            \c pData parameter.
	/// \param[in] pData The data to transfer.
	/// \param[in] takeStgMediumOwnership If \c TRUE the data object is responsible for freeing the storage
	///            medium defined by the \c pData parameter; otherwise this is done by the caller of this
	///            method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, EnumFormatEtc, DAdvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms686626.aspx">IDataObject::SetData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683812.aspx">STGMEDIUM</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693491.aspx">ReleaseStgMedium</a>
	virtual HRESULT STDMETHODCALLTYPE SetData(FORMATETC* pFormat, STGMEDIUM* pData, BOOL takeStgMediumOwnership);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IEnumFORMATETC
	///
	//@{
	/// \brief <em>Clones the \c FORMATETC iterator used to iterate the data format</em>
	///
	/// Clones the \c FORMATETC iterator including its current state. This iterator is used to iterate
	/// the data format descriptors offered by this object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumFORMATETC implementation.
	///
	/// \return This object's new reference count.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682337.aspx">IEnumFORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumFORMATETC** /*ppEnumerator*/);
	/// \brief <em>Retrieves the next x data formats</em>
	///
	/// Retrieves the next \c numberOfMaxFormats data format descriptors from the iterator.
	///
	/// \param[in] numberOfMaxFormats The maximum number of data format descriptors the array identified by
	///            \c pFormats can contain.
	/// \param[in,out] pFormats An array of \c FORMATETC structs. On return, each \c FORMATETC will contain
	///                the description of a data format offered by this object.
	/// \param[out] pNumberOfFormatsReturned The number of data format descriptors that actually were copied
	///             to the array identified by \c pFormats.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxFormats, FORMATETC* pFormats, ULONG* pNumberOfFormatsReturned);
	/// \brief <em>Resets the \c FORMATETC iterator</em>
	///
	/// Resets the \c FORMATETC iterator so that the next call of \c Next or \c Skip starts at the first data
	/// format descriptor.
	///
	/// \return The object's new reference count.
	///
	/// \sa Clone, Next, Skip
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x data formats</em>
	///
	/// Instructs the \c FORMATETC iterator to skip the next \c numberOfFormatsToSkip data format
	/// descriptors.
	///
	/// \param[in] numberOfFormatsToSkip The number of data format descriptors to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfFormatsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds a detailed data format description and the corresponding data</em>
	typedef struct DataEntry
	{
		/// \brief <em>The detailed data format description</em>
		LPFORMATETC pFormat;
		/// \brief <em>The data</em>
		LPSTGMEDIUM pData;

		/// \brief <em>Frees the specified \c DataEntry memory</em>
		///
		/// \param[in] pEntry The \c DataEntry struct to free. It must have been allocated using \c HeapAlloc.
		///
		/// \remarks \c pEntry will be invalid when the function returns.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/aa366597.aspx">HeapAlloc</a>
		static void Release(__in DataEntry* pEntry)
		{
			if(pEntry->pFormat) {
				HeapFree(GetProcessHeap(), 0, pEntry->pFormat);
			}
			if(pEntry->pData) {
				ReleaseStgMedium(pEntry->pData);
				HeapFree(GetProcessHeap(), 0, pEntry->pData);
			}
			HeapFree(GetProcessHeap(), 0, pEntry);
		}
	} DataEntry;

	/// \brief <em>Holds all detailed data format descriptions and the corresponding data for a specific data format</em>
	typedef struct DataFormat
	{
		#ifdef USE_STL
			/// \brief <em>Holds any data that matches the format, this struct is used for</em>
			std::vector<DataEntry*> dataEntries;
		#else
			/// \brief <em>Holds any data that matches the format, this struct is used for</em>
			CAtlArray<DataEntry*> dataEntries;
		#endif

		/// \brief <em>Frees the specified \c DataFormat memory</em>
		///
		/// \param[in] pFormat The \c DataFormat struct to free. It must have been allocated using the \c new
		///            keyword.
		///
		/// \remarks \c pFormat will be invalid when the function returns.
		static void Release(__in DataFormat* pFormat)
		{
			#ifdef USE_STL
				for(std::vector<DataEntry*>::iterator iter = pFormat->dataEntries.begin(); iter != pFormat->dataEntries.end(); ++iter) {
					ATLASSERT_POINTER(*iter, DataEntry);
					DataEntry::Release(*iter);
				}
				pFormat->dataEntries.clear();
			#else
				for(size_t i = 0; i < pFormat->dataEntries.GetCount(); ++i) {
					ATLASSERT_POINTER(pFormat->dataEntries[i], DataEntry);
					DataEntry::Release(pFormat->dataEntries[i]);
				}
				pFormat->dataEntries.RemoveAll();
			#endif
			delete pFormat;
		}
	} DataFormat;

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c TextBox object that owns this object</em>
		///
		/// \sa SetOwner
		TextBox* pOwnerTxtBox;
		/// \brief <em>Holds the data advise holder object</em>
		LPDATAADVISEHOLDER pDataAdviseHolder;
		#ifdef USE_STL
			/// \brief <em>Holds any data this object provides</em>
			std::unordered_map<CLIPFORMAT, DataFormat*> containedData;
			/// \brief <em>Points to the next enumerated contained data</em>
			std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator nextEnumeratedData;
		#else
			/// \brief <em>Holds any data this object provides</em>
			CAtlMap<CLIPFORMAT, DataFormat*> containedData;
			/// \brief <em>Points to the next enumerated contained data</em>
			POSITION nextEnumeratedData;
		#endif
		/// \brief <em>Holds the number of dragged items</em>
		///
		/// We'll add a label displaying the number of dragged items to the drag image if the number is higher
		/// than 1.
		///
		/// \sa SetData
		LONG numberOfItemsToDisplay;

		Properties()
		{
			pOwnerTxtBox = NULL;
			pDataAdviseHolder = NULL;
			#ifdef USE_STL
				nextEnumeratedData = containedData.begin();
			#else
				nextEnumeratedData = containedData.GetStartPosition();
			#endif
			numberOfItemsToDisplay = 0;
		}

		~Properties();
	} properties;
};     // SourceOLEDataObject

OBJECT_ENTRY_AUTO(__uuidof(OLEDataObject), SourceOLEDataObject)