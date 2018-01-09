//////////////////////////////////////////////////////////////////////
/// \class VirtualUpDownAccelerator
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing up down accelerator</em>
///
/// This class is a wrapper around an up down accelerator that does not yet exist in the up down control.
///
/// \if UNICODE
///   \sa EditCtlsLibU::IVirtualUpDownAccelerator, UpDownAccelerator, VirtualUpDownAccelerators,
///       UpDownTextBox
/// \else
///   \sa EditCtlsLibA::IVirtualUpDownAccelerator, UpDownAccelerator, VirtualUpDownAccelerators,
///       UpDownTextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "EditCtlsU.h"
#else
	#include "EditCtlsA.h"
#endif
#include "_IVirtualUpDownAcceleratorEvents_CP.h"
#include "helpers.h"
#include "UpDownTextBox.h"


class ATL_NO_VTABLE VirtualUpDownAccelerator : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualUpDownAccelerator, &CLSID_VirtualUpDownAccelerator>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualUpDownAccelerator>,
    public Proxy_IVirtualUpDownAcceleratorEvents<VirtualUpDownAccelerator>,
    #ifdef _UNICODE
    	public IDispatchImpl<IVirtualUpDownAccelerator, &IID_IVirtualUpDownAccelerator, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualUpDownAccelerator, &IID_IVirtualUpDownAccelerator, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALUPDOWNACCELERATOR)

		BEGIN_COM_MAP(VirtualUpDownAccelerator)
			COM_INTERFACE_ENTRY(IVirtualUpDownAccelerator)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualUpDownAccelerator)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualUpDownAcceleratorEvents))
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
	/// \name Implementation of IVirtualUpDownAccelerator
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
	/// \remarks This property is read-only.
	///
	/// \sa get_StepSize, get_Index
	virtual HRESULT STDMETHODCALLTYPE get_ActivationTime(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this accelerator.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c StepSize property</em>
	///
	/// Retrieves the increment by which this accelerator changes the up down control's value.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IUpDownAccelerator interface.\n
	///          This property is read-only.
	///
	/// \sa get_ActivationTime, UpDownTextBox::get_CurrentValue, get_Index
	virtual HRESULT STDMETHODCALLTYPE get_StepSize(LONG* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given accelerator</em>
	///
	/// Attaches this object to a given accelerator, so that the accelerator's properties can be
	/// retrieved and set using this object's methods.
	///
	/// \param[in] pAccelerators The control's accelerators.
	/// \param[in] acceleratorIndex The accelerator to attach to.
	///
	/// \sa Detach
	void Attach(LPUDACCEL pAccelerators, int acceleratorIndex);
	/// \brief <em>Detaches this object from a accelerator</em>
	///
	/// Detaches this object from the accelerator it currently wraps, so that it doesn't wrap any
	/// accelerator anymore.
	///
	/// \sa Attach
	void Detach(void);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the control's accelerators</em>
		LPUDACCEL pAccelerators;
		/// \brief <em>The index of the accelerator wrapped by this object</em>
		int acceleratorIndex;

		Properties()
		{
			acceleratorIndex = -1;
			pAccelerators = NULL;
		}
	} properties;
};     // VirtualUpDownAccelerator

OBJECT_ENTRY_AUTO(__uuidof(VirtualUpDownAccelerator), VirtualUpDownAccelerator)