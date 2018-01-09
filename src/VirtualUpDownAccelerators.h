//////////////////////////////////////////////////////////////////////
/// \class VirtualUpDownAccelerators
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c VirtualUpDownAccelerator objects</em>
///
/// This class provides easy access to collections of \c VirtualUpDownAccelerator objects. A
/// \c VirtualUpDownAccelerators object is used to group the up down accelerators that do not yet exist in
/// the up down control.
///
/// \if UNICODE
///   \sa EditCtlsLibU::IVirtualUpDownAccelerators, VirtualUpDownAccelerator, UpDownAccelerators,
///       UpDownTextBox
/// \else
///   \sa EditCtlsLibA::IVirtualUpDownAccelerators, VirtualUpDownAccelerator, UpDownAccelerators,
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
#include "_IVirtualUpDownAcceleratorsEvents_CP.h"
#include "helpers.h"
#include "UpDownTextBox.h"
#include "VirtualUpDownAccelerator.h"


class ATL_NO_VTABLE VirtualUpDownAccelerators : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualUpDownAccelerators, &CLSID_VirtualUpDownAccelerators>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualUpDownAccelerators>,
    public Proxy_IVirtualUpDownAcceleratorsEvents<VirtualUpDownAccelerators>,
    public IEnumVARIANT,
    #ifdef _UNICODE
    	public IDispatchImpl<IVirtualUpDownAccelerators, &IID_IVirtualUpDownAccelerators, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualUpDownAccelerators, &IID_IVirtualUpDownAccelerators, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALUPDOWNACCELERATORS)

		BEGIN_COM_MAP(VirtualUpDownAccelerators)
			COM_INTERFACE_ENTRY(IVirtualUpDownAccelerators)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualUpDownAccelerators)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualUpDownAcceleratorsEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the accelerators</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c VirtualUpDownAccelerator objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x accelerators</em>
	///
	/// Retrieves the next \c numberOfMaxItems accelerators from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of accelerators the array identified by \c pItems
	///            can contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to an accelerator's \c IVirtualUpDownAccelerator implementation.
	/// \param[out] pNumberOfItemsReturned The number of accelerators that actually were copied to the
	///             array identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, VirtualUpDownAccelerator,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// accelerator in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x accelerators</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip accelerators.
	///
	/// \param[in] numberOfItemsToSkip The number of accelerators to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IVirtualUpDownAccelerators
	///
	//@{
	/// \brief <em>Retrieves a \c VirtualUpDownAccelerator object from the collection</em>
	///
	/// Retrieves a \c VirtualUpDownAccelerator object from the collection that wraps the up down accelerator
	/// identified by \c acceleratorIndex.
	///
	/// \param[in] acceleratorIndex A value that identifies the up down accelerator to be retrieved.
	/// \param[out] ppAccelerator Receives the accelerator's \c IVirtualUpDownAccelerator implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa VirtualUpDownAccelerator
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG acceleratorIndex, IVirtualUpDownAccelerator** ppAccelerator);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c VirtualUpDownAccelerator objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Counts the accelerators in the collection</em>
	///
	/// Retrieves the number of \c VirtualUpDownAccelerator objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given accelerator collection</em>
	///
	/// Attaches this object to a given accelerator collection, so that the collection can be managed
	/// using this object's methods.
	///
	/// \param[in] numberOfAccelerators The number of accelerators defined by the \c pAccelerators parameter.
	/// \param[in] pAccelerators The accelerators to attach to.
	///
	/// \sa Detach
	void Attach(int numberOfAccelerators, LPUDACCEL pAccelerators);
	/// \brief <em>Detaches this object from an accelerator collection</em>
	///
	/// Detaches this object from the accelerator collection it currently wraps, so that it doesn't wrap
	/// any accelerator collection anymore.
	///
	/// \sa Attach
	void Detach(void);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the number of accelerators wrapped by this collection</em>
		int numberOfAccelerators;
		/// \brief <em>Holds the accelerators wrapped by this collection</em>
		LPUDACCEL pAccelerators;
		/// \brief <em>Holds the next enumerated accelerator</em>
		int nextEnumeratedAccelerator;

		Properties()
		{
			numberOfAccelerators = 0;
			pAccelerators = NULL;
			nextEnumeratedAccelerator = 0;
		}
	} properties;
};     // VirtualUpDownAccelerators

OBJECT_ENTRY_AUTO(__uuidof(VirtualUpDownAccelerators), VirtualUpDownAccelerators)