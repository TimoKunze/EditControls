//////////////////////////////////////////////////////////////////////
/// \class UpDownAccelerators
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c UpDownAccelerator objects</em>
///
/// This class provides easy access (including filtering) to collections of \c UpDownAccelerator objects.
/// \c UpDownAccelerators objects are used to group up down accelerators that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa EditCtlsLibU::IUpDownAccelerators, UpDownAccelerator, UpDownTextBox
/// \else
///   \sa EditCtlsLibA::IUpDownAccelerators, UpDownAccelerator, UpDownTextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "EditCtlsU.h"
#else
	#include "EditCtlsA.h"
#endif
#include "_IUpDownAcceleratorsEvents_CP.h"
#include "helpers.h"
#include "UpDownTextBox.h"
#include "UpDownAccelerator.h"


class ATL_NO_VTABLE UpDownAccelerators : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<UpDownAccelerators, &CLSID_UpDownAccelerators>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<UpDownAccelerators>,
    public Proxy_IUpDownAcceleratorsEvents<UpDownAccelerators>,
    public IEnumVARIANT,
    #ifdef _UNICODE
    	public IDispatchImpl<IUpDownAccelerators, &IID_IUpDownAccelerators, &LIBID_EditCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IUpDownAccelerators, &IID_IUpDownAccelerators, &LIBID_EditCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_UPDOWNACCELERATORS)

		BEGIN_COM_MAP(UpDownAccelerators)
			COM_INTERFACE_ENTRY(IUpDownAccelerators)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(UpDownAccelerators)
			CONNECTION_POINT_ENTRY(__uuidof(_IUpDownAcceleratorsEvents))
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
	/// the \c UpDownAccelerator objects managed by this collection object.
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
	/// Retrieves the next \c numberOfMaxAccelerators accelerators from the iterator.
	///
	/// \param[in] numberOfMaxAccelerators The maximum number of accelerators the array identified by
	///            \c pAccelerators can contain.
	/// \param[in,out] pAccelerators An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a accelerator's \c IUpDownAccelerator implementation.
	/// \param[out] pNumberOfAcceleratorsReturned The number of accelerators that actually were copied to the
	///             array identified by \c pAccelerators.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, UpDownAccelerator,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxAccelerators, VARIANT* pAccelerators, ULONG* pNumberOfAcceleratorsReturned);
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
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfAcceleratorsToSkip items.
	///
	/// \param[in] numberOfAcceleratorsToSkip The number of accelerators to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfAcceleratorsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IUpDownAccelerators
	///
	//@{
	/// \brief <em>Retrieves a \c UpDownAccelerator object from the collection</em>
	///
	/// Retrieves a \c UpDownAccelerator object from the collection that wraps the up down accelerator
	/// identified by \c acceleratorIndex.
	///
	/// \param[in] acceleratorIndex A value that identifies the up down accelerator to be retrieved.
	/// \param[out] ppAccelerator Receives the accelerator's \c IUpDownAccelerator implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa UpDownAccelerator, Add, Remove
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG acceleratorIndex, IUpDownAccelerator** ppAccelerator);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c UpDownAccelerator objects
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

	/// \brief <em>Adds an accelerator to the up down control</em>
	///
	/// Adds an accelerator with the specified properties to the control and returns an \c UpDownAccelerator
	/// object wrapping the inserted accelerator.
	///
	/// \param[in] activationTime The number of seconds that must elapse, before the new accelerator is
	///            used.
	/// \param[in] stepSize The increment by which the new accelerator changes the up down control's value.
	/// \param[out] ppAddedAccelerator Receives the added accelerator's \c IUpDownAccelerator implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The accelerators will be sorted by ascending activation time and ascending step size.
	///
	/// \sa Count, Remove, RemoveAll, UpDownAccelerator::get_ActivationTime, UpDownAccelerator::get_StepSize,
	///     UpDownTextBox::get_CurrentValue
	virtual HRESULT STDMETHODCALLTYPE Add(LONG activationTime, LONG stepSize, IUpDownAccelerator** ppAddedAccelerator = NULL);
	/// \brief <em>Counts the accelerators in the collection</em>
	///
	/// Retrieves the number of \c UpDownAccelerator objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified accelerator in the collection from the up down control</em>
	///
	/// \param[in] acceleratorIndex A value that identifies the up down accelerator to be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks An up down control must have at least one accelerator. If calling this method would make
	///          this condition fail, a new accelerator with \c ActivationTime and \c StepSize both set to 0
	///          is inserted.
	///
	/// \sa Add, Count, RemoveAll, UpDownAccelerator::get_ActivationTime, UpDownAccelerator::get_StepSize
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG acceleratorIndex);
	/// \brief <em>Removes all accelerators in the collection from the up down control</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks An up down control must have at least one accelerator. If calling this method would make
	///          this condition fail, a new accelerator with \c ActivationTime and \c StepSize both set to 0
	///          is inserted.
	///
	/// \sa Add, Count, Remove, UpDownAccelerator::get_ActivationTime, UpDownAccelerator::get_StepSize
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerUpDownTextBox
	void SetOwner(__in_opt UpDownTextBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c UpDownTextBox object that owns this collection</em>
		///
		/// \sa SetOwner
		UpDownTextBox* pOwnerUpDownTextBox;
		/// \brief <em>Holds the next enumerated accelerator</em>
		int nextEnumeratedAccelerator;

		Properties()
		{
			pOwnerUpDownTextBox = NULL;
			nextEnumeratedAccelerator = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning up down control's window handle</em>
		///
		/// \return The window handle of the up down control that contains the accelerators in this collection.
		///
		/// \sa pOwnerUpDownTextBox
		HWND GetUpDownHWnd(void);
	} properties;
};     // UpDownAccelerators

OBJECT_ENTRY_AUTO(__uuidof(UpDownAccelerators), UpDownAccelerators)