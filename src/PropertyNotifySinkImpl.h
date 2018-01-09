//////////////////////////////////////////////////////////////////////
/// \class PropertyNotifySinkImpl
/// \author Timo "TimoSoft" Kunze
/// \brief <em>An implementation of the \c IPropertyNotifySink interface</em>
///
/// This class implements the \c IPropertyNotifySink interface and can be used to receive a
/// notification if an object's property is changed.
/// We use this for Font and Picture properties.
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa HotKeyBox, IPAddressBox, TextBox, UpDownTextBox,
///     <a href="https://msdn.microsoft.com/en-us/library/ms692638.aspx">IPropertyNotifySink</a>
//////////////////////////////////////////////////////////////////////

#pragma once


template <class TypeOfOwner>
class ATL_NO_VTABLE PropertyNotifySinkImpl :
    public CComObjectRootEx<CComSingleThreadModel>,
    public IPropertyNotifySink
{
public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_COM_MAP(PropertyNotifySinkImpl)
			COM_INTERFACE_ENTRY(IPropertyNotifySink)
		END_COM_MAP()
	#endif


	/// \brief <em>Initializes the object</em>
	///
	/// Binds the object to a property object and stores a pointer to the object to notify on changes.
	///
	/// \param[in] pObjectToNotify The object to notify on changes.
	/// \param[in] propertyToWatch The property to watch for changes.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa StartWatching
	HRESULT Initialize(TypeOfOwner* pObjectToNotify, DISPID propertyToWatch)
	{
		if(!pObjectToNotify) {
			return E_FAIL;
		}

		properties.pObjectToNotify = pObjectToNotify;
		properties.propertyToWatch = propertyToWatch;
		return S_OK;
	}
	/// \brief <em>Starts watching the property object for changes</em>
	///
	/// \param[in] pPropertyObject The property object to watch.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa StopWatching, Initialize
	HRESULT StartWatching(IUnknown* pPropertyObject)
	{
		if(!pPropertyObject) {
			return E_FAIL;
		}
		return AtlAdvise(pPropertyObject, this, IID_IPropertyNotifySink, &properties.notifyCookie);
	}
	/// \brief <em>Stops watching the property object for changes</em>
	///
	/// \param[in] pPropertyObject The property object to stop watching.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa StartWatching, Initialize
	HRESULT StopWatching(IUnknown* pPropertyObject)
	{
		if(!pPropertyObject) {
			return E_FAIL;
		}
		return AtlUnadvise(pPropertyObject, IID_IPropertyNotifySink, properties.notifyCookie);
	}

	/// \brief <em>Will be called after property changes</em>
	///
	/// Will be called after the property object's property with the DISPID \c changedProperty was changed.
	///
	/// \param[in] changedProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	HRESULT STDMETHODCALLTYPE OnChanged(DISPID changedProperty)
	{
		return properties.pObjectToNotify->OnPropertyObjectChanged(properties.propertyToWatch, changedProperty);
	}
	/// \brief <em>Will be called before property changes</em>
	///
	/// Will be called if the property object's property with the DISPID \c requestedProperty is about to be
	/// changed.
	///
	/// \param[in] requestedProperty The \c DISPID of the property that is about to change.
	///
	/// \return An \c HRESULT error code.
	HRESULT STDMETHODCALLTYPE OnRequestEdit(DISPID requestedProperty)
	{
		return properties.pObjectToNotify->OnPropertyObjectRequestEdit(properties.propertyToWatch, requestedProperty);
	}

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The object to notify on changes</em>
		TypeOfOwner* pObjectToNotify;
		/// \brief <em>The cookie used with the \c AtlAdvise() and \c AtlUnadvise() methods</em>
		DWORD notifyCookie;
		/// \brief <em>Holds the ID of the property we're watching for changes</em>
		DISPID propertyToWatch;

		Properties()
		{
			pObjectToNotify = NULL;
			notifyCookie = 0;
			propertyToWatch = 0;
		}
	} properties;
};     // PropertyNotifySinkImpl