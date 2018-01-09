//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa HotKeyBox, IPAddressBox, TextBox, UpDownTextBox
//////////////////////////////////////////////////////////////////////


#pragma once

#include "UpDownAccelerator.h"
#include "UpDownAccelerators.h"
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	/// \brief <em>Creates a \c UpDownAccelerator object</em>
	///
	/// Creates a \c UpDownAccelerator object that represents a given up down accelerator and returns
	/// its \c IUpDownAccelerator implementation as a smart pointer.
	///
	/// \param[in] acceleratorIndex The index of the accelerator for which to create the object.
	/// \param[in] pUpDownTextBox The \c UpDownTextBox object the \c UpDownAccelerator object will work on.
	/// \param[in] validateAcceleratorIndex If \c TRUE, the method will check whether the \c acceleratorIndex
	///            parameter identifies an existing accelerator; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IUpDownAccelerator implementation.
	static inline CComPtr<IUpDownAccelerator> InitAccelerator(int acceleratorIndex, UpDownTextBox* pUpDownTextBox, BOOL validateAcceleratorIndex = TRUE)
	{
		CComPtr<IUpDownAccelerator> pAccelerator = NULL;
		InitAccelerator(acceleratorIndex, pUpDownTextBox, IID_IUpDownAccelerator, reinterpret_cast<LPUNKNOWN*>(&pAccelerator), validateAcceleratorIndex);
		return pAccelerator;
	};

	/// \brief <em>Creates a \c UpDownAccelerator object</em>
	///
	/// \overload
	///
	/// Creates a \c UpDownAccelerator object that represents a given up down accelerator and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] acceleratorIndex The index of the accelerator for which to create the object.
	/// \param[in] pUpDownTextBox The \c UpDownTextBox object the \c UpDownAccelerator object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppAccelerator Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateAcceleratorIndex If \c TRUE, the method will check whether the \c acceleratorIndex
	///            parameter identifies an existing accelerator; otherwise not.
	static inline void InitAccelerator(int acceleratorIndex, UpDownTextBox* pUpDownTextBox, REFIID requiredInterface, LPUNKNOWN* ppAccelerator, BOOL validateAcceleratorIndex = TRUE)
	{
		ATLASSERT_POINTER(ppAccelerator, LPUNKNOWN);
		ATLASSUME(ppAccelerator);

		*ppAccelerator = NULL;
		if(validateAcceleratorIndex && !IsValidAcceleratorIndex(acceleratorIndex, pUpDownTextBox)) {
			// there's nothing useful we could return
			return;
		}

		// create a UpDownAccelerator object and initialize it with the given parameters
		CComObject<UpDownAccelerator>* pUpDownAccelObj = NULL;
		CComObject<UpDownAccelerator>::CreateInstance(&pUpDownAccelObj);
		pUpDownAccelObj->AddRef();
		pUpDownAccelObj->SetOwner(pUpDownTextBox);
		pUpDownAccelObj->Attach(acceleratorIndex);
		pUpDownAccelObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppAccelerator));
		pUpDownAccelObj->Release();
	};

	/// \brief <em>Creates a \c UpDownAccelerators object</em>
	///
	/// Creates a \c UpDownAccelerators object that represents a collection of up down accelerators and
	/// returns its \c IUpDownAccelerators implementation as a smart pointer.
	///
	/// \param[in] pUpDownTextBox The \c UpDownTextBox object the \c UpDownAccelerators object will work on.
	///
	/// \return A smart pointer to the created object's \c IUpDownAccelerators implementation.
	static inline CComPtr<IUpDownAccelerators> InitAccelerators(UpDownTextBox* pUpDownTextBox)
	{
		CComPtr<IUpDownAccelerators> pAccelerators = NULL;
		InitAccelerators(pUpDownTextBox, IID_IUpDownAccelerators, reinterpret_cast<LPUNKNOWN*>(&pAccelerators));
		return pAccelerators;
	};

	/// \brief <em>Creates a \c UpDownAccelerators object</em>
	///
	/// \overload
	///
	/// Creates a \c UpDownAccelerators object that represents a collection of up down accelerators and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pUpDownTextBox The \c UpDownTextBox object the \c UpDownAccelerators object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppAccelerators Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitAccelerators(UpDownTextBox* pUpDownTextBox, REFIID requiredInterface, LPUNKNOWN* ppAccelerators)
	{
		ATLASSERT_POINTER(ppAccelerators, LPUNKNOWN);
		ATLASSUME(ppAccelerators);

		*ppAccelerators = NULL;
		// create a UpDownAccelerators object and initialize it with the given parameters
		CComObject<UpDownAccelerators>* pUpDownAccelsObj = NULL;
		CComObject<UpDownAccelerators>::CreateInstance(&pUpDownAccelsObj);
		pUpDownAccelsObj->AddRef();
		pUpDownAccelsObj->SetOwner(pUpDownTextBox);
		pUpDownAccelsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppAccelerators));
		pUpDownAccelsObj->Release();
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};
};     // ClassFactory