//////////////////////////////////////////////////////////////////////
/// \class EnumOLEVERB
/// \author Timo "TimoSoft" Kunze
/// \brief <em>An implementation of the \c IEnumOLEVERB interface</em>
///
/// This class implements \c IEnumOLEVERB and can be used to provide OLE verbs for an
/// ActiveX control.
///
/// \todo Improve documentation of the \c pOuter parameter in \c CreateInstance().
/// \todo Improve documentation.
///
/// \sa HotKeyBox, IPAddressBox, TextBox, UpDownTextBox,
///     <a href="https://msdn.microsoft.com/en-us/library/ms695084.aspx">IEnumOLEVERB</a>
//////////////////////////////////////////////////////////////////////

#pragma once


class EnumOLEVERB :
    public IEnumOLEVERB
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// \param[in] pVerbs An array of OLE verbs.
	/// \param[in] numberOfVerbs The number of elements in the array identified by \c pVerbs.
	///
	/// \sa ~EnumOLEVERB
	EnumOLEVERB(const OLEVERB* pVerbs, int numberOfVerbs);
	/// \brief <em>The destructor of this class</em>
	///
	/// Does nothing at the moment.
	///
	/// \sa EnumOLEVERB
	~EnumOLEVERB();

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IUnknown
	///
	//@{
	/// \brief <em>Adds a reference to this object</em>
	///
	/// Increases the references counter of this object by 1.
	///
	/// \return The new reference count.
	///
	/// \sa Release, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
	ULONG STDMETHODCALLTYPE AddRef();
	/// \brief <em>Removes a reference from this object</em>
	///
	/// Decreases the references counter of this object by 1. If the counter reaches 0, the object
	/// is destroyed.
	///
	/// \return The new reference count.
	///
	/// \sa AddRef, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
	ULONG STDMETHODCALLTYPE Release();
	/// \brief <em>Queries for an interface implementation</em>
	///
	/// Queries this object for its implementation of a given interface.
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AddRef, Release,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IEnumOLEVERB
	///
	//@{
	/// \brief <em>Clones the OLE verb iterator</em>
	///
	/// Clones the OLE verb iterator including its current state.
	///
	/// \param[out] ppEnumerator Receives the clone's implementation of the \c IEnumOLEVERB interface.
	///
	/// \return The new reference count.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695084.aspx">IEnumOLEVERB</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	HRESULT STDMETHODCALLTYPE Clone(IEnumOLEVERB** ppEnumerator);
	/// \brief <em>Retrieves the next x OLE verbs</em>
	///
	/// Retrieves the next \c numberOfMaxVerbs OLE verbs from the iterator.
	///
	/// \param[in] numberOfMaxVerbs The maximum number of verbs the array identified by \c pVerbs
	///            can contain.
	/// \param[out] pVerbs An array of OLE verbs. Must be large enough to hold the number of elements
	///             specified by \c numberOfMaxVerbs.
	/// \param[out] pNumberOfVerbsReturned The number of verbs that actually were copied into the array
	///             identified by \c pVerbs.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxVerbs, OLEVERB* pVerbs, ULONG* pNumberOfVerbsReturned);
	/// \brief <em>Resets the OLE verb iterator</em>
	///
	/// Resets the OLE verb iterator so that the next call of \c Next() or \c Skip() starts at the first
	/// OLE verb in the collection.
	///
	/// \return The new reference count.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	HRESULT STDMETHODCALLTYPE Reset();
	/// \brief <em>Skips the next x OLE verbs</em>
	///
	/// Instructs the OLE verb iterator to skip the next \c numberOfVerbsToSkip verbs.
	///
	/// \param[in] numberOfVerbsToSkip The number of OLE verbs to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfVerbsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Creates an instance of this class</em>
	///
	/// Creates an instance of this class, initializes the object with the given values and returns
	/// the implementation of a given interface.
	///
	/// \param[in] pVerbs An array of OLE verbs.
	/// \param[in] numberOfVerbs The number of elements in the array identified by \c pVerbs.
	/// \param[in] pOuter The outer unknown or controlling unknown of the aggregate.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	static HRESULT CreateInstance(const OLEVERB* pVerbs, int numberOfVerbs, LPUNKNOWN pOuter, REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	/// \brief <em>Creates an instance of this class</em>
	///
	/// Creates an instance of this class, initializes the object with the given values and returns
	/// the \c IEnumOLEVERB implementation.
	///
	/// \param[in] pVerbs An array of OLE verbs.
	/// \param[in] numberOfVerbs The number of elements in the array identified by \c pVerbs.
	/// \param[out] ppEnumerator Receives the object's implementation of the \c IEnumOLEVERB interface.
	///
	/// \return An \c HRESULT error code.
	static HRESULT CreateInstance(const OLEVERB* pVerbs, int numberOfVerbs, IEnumOLEVERB** ppEnumerator);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the object's reference count</em>
		LONG referenceCount;
		/// \brief <em>Holds the index of the next OLE verb to retrieve during enumeration</em>
		int nextVerb;
		/// \brief <em>Holds a pointer to the array of all OLE verbs</em>
		const OLEVERB* pVerbs;
		/// \brief <em>Holds the number of all OLE verbs</em>
		int numberOfVerbs;

		Properties()
		{
			referenceCount = 0;
			nextVerb = 0;
			numberOfVerbs = 0;
			pVerbs = NULL;
		}
	} properties;
};     // EnumOLEVERB