//////////////////////////////////////////////////////////////////////
/// \class ICategorizeProperties
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for categorizing COM properties</em>
///
/// This interface can be implemented by COM objects to categorize their properties.\n
/// It was written by Microsoft, but never made it into any official header file.
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa HotKeyBox, IPAddressBox, TextBox, UpDownTextBox
//////////////////////////////////////////////////////////////////////


#pragma once


// the interface's GUID
const IID IID_ICategorizeProperties = {0x4D07FC10, 0xF931, 0x11CE, {0xB0, 0x01, 0x00, 0xAA, 0x00, 0x68, 0x84, 0xE5}};
typedef int PROPCAT;

// predefined categories
const int PROPCAT_Nil					= -1;
const int PROPCAT_Misc				= -2;
const int PROPCAT_Font				= -3;
const int PROPCAT_Position		= -4;
const int PROPCAT_Appearance	= -5;
const int PROPCAT_Behavior		= -6;
const int PROPCAT_Data				= -7;
const int PROPCAT_List				= -8;
const int PROPCAT_Text				= -9;
const int PROPCAT_Scale				= -10;
const int PROPCAT_DDE					= -11;

// user-defined property categories
const int PROPCAT_Colors			= 1;
const int PROPCAT_DragDrop		= 2;


class ICategorizeProperties :
    public IUnknown
{
public:
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory) = 0;
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	/// \param[in] languageID The locale identifier of the language the name should be in.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID languageID, BSTR* pName) = 0;
};