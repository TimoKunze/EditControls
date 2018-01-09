//////////////////////////////////////////////////////////////////////
/// \class UpDownAccelerator
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing up down accelerator</em>
///
/// This class is a wrapper around an up down accelerator that - unlike a working area wrapped by
/// \c VirtualUpDownAccelerator - really exists in the up down control.
///
/// \if UNICODE
///   \sa EditCtlsLibU::IUpDownAccelerator, VirtualUpDownAccelerator, UpDownAccelerators, UpDownTextBox
/// \else
///   \sa EditCtlsLibA::IUpDownAccelerator, VirtualUpDownAccelerator, UpDownAccelerators, UpDownTextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "EditCtlsU.h"
#else
	#include "EditCtlsA.h"
#endif
#include "_IUpDownAcceleratorEvents_CP.h"
#include "helpers.h"
#include "UpDownTextBox.h"


class ATL_NO_VTABLE UpDownAccelerator : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<UpDownAccelerator, &CLSID_UpDownAccelerator>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<UpDownAccelerator>,
    public Proxy_IUpDownAcceleratorEvents<UpDownAccelerator>, 
    #ifdef _UNICODE
    	public IDispatchImpl<IUpDownAccelerator, &IID_IUpDownAccelerator, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IUpDownAccelerator, &IID_IUpDownAccelerator, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_UPDOWNACCELERATOR)

		BEGIN_COM_MAP(UpDownAccelerator)
			COM_INTERFACE_ENTRY(IUpDownAccelerator)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(UpDownAccelerator)
			CONNECTION_POINT_ENTRY(__uuidof(_IUpDownAcceleratorEvents))
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
	/// \name Implementation of IUpDownAccelerator
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c ActivationTime property</em>
	///
	/// Retrieves the number of seconds that must elapse, before this accelerator is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property may change the \c Index property, because the accelerators will be
	///          sorted again.
	///
	/// \sa put_ActivationTime, get_StepSize, get_Index
	virtual HRESULT STDMETHODCALLTYPE get_ActivationTime(LONG* pValue);
	/// \brief <em>Sets the \c ActivationTime property</em>
	///
	/// Sets the number of seconds that must elapse, before this accelerator is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property may change the \c Index property, because the accelerators will be
	///          sorted again.
	///
	/// \sa get_ActivationTime, put_StepSize, get_Index
	virtual HRESULT STDMETHODCALLTYPE put_ActivationTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this accelerator.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The accelerators are sorted by activation time and step size. Therefore, adding or
	///          removing accelerators or changing an accelerator's properties may change other
	///          accelerators' indexes.\n
	///          This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c StepSize property</em>
	///
	/// Retrieves the increment by which this accelerator changes the up down control's value.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property may change the \c Index property, because the accelerators will be
	///          sorted again.\n
	///          This is the default property of the \c IUpDownAccelerator interface.
	///
	/// \sa put_StepSize, get_ActivationTime, UpDownTextBox::get_CurrentValue, get_Index
	virtual HRESULT STDMETHODCALLTYPE get_StepSize(LONG* pValue);
	/// \brief <em>Sets the \c StepSize property</em>
	///
	/// Sets the increment by which this accelerator changes the up down control's value.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property may change the \c Index property, because the accelerators will be
	///          sorted again.\n
	///          This is the default property of the \c IUpDownAccelerator interface.
	///
	/// \sa get_StepSize, put_ActivationTime, UpDownTextBox::put_CurrentValue, get_Index
	virtual HRESULT STDMETHODCALLTYPE put_StepSize(LONG newValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given accelerator</em>
	///
	/// Attaches this object to a given accelerator, so that the accelerator's properties can be
	/// retrieved and set using this object's methods.
	///
	/// \param[in] acceleratorIndex The accelerator to attach to.
	///
	/// \sa Detach
	void Attach(int acceleratorIndex);
	/// \brief <em>Detaches this object from an accelerator</em>
	///
	/// Detaches this object from the accelerator it currently wraps, so that it doesn't wrap any
	/// accelerator anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this accelerator</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerUpDownTextBox
	void SetOwner(__in_opt UpDownTextBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c UpDownTextBox object that owns this accelerator</em>
		///
		/// \sa SetOwner
		UpDownTextBox* pOwnerUpDownTextBox;
		/// \brief <em>The index of the accelerator wrapped by this object</em>
		int acceleratorIndex;

		Properties()
		{
			pOwnerUpDownTextBox = NULL;
			acceleratorIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning up down control's window handle</em>
		///
		/// \return The window handle of the up down control that contains this accelerator.
		///
		/// \sa pOwnerUpDownTextBox
		HWND GetUpDownHWnd(void);
	} properties;
};     // UpDownAccelerator

OBJECT_ENTRY_AUTO(__uuidof(UpDownAccelerator), UpDownAccelerator)