#pragma once
#include <afxtempl.h>

/*
* Classes to translate tooltip text requests received at the main frame to the corresponding elements (buttons) owners.
*/

/*

The problem is that the custom tooltip text messages are sent not to the controls the mouse is howering over,
but to the top-level main frame of the application.
To customize tooltip texts at the related control level, we do the following.
We are hooking the "get tooltip text" message at the main frame level.
There is a function for this already, CMDIFrameWndEx::GetToolbarButtonToolTipText.
Then we route the message to the related control (provided the control is subcribed for this message processing).

Usage:

- Instantiate CToolTipTextProviderList in the main frame context.
- Override CMDIFrameWndEx::GetToolbarButtonToolTipText at the main frame class, translate the call to the CToolTipTextProviderList::OnGetToolTipText.

	class MainFrame : public CMDIFrameWndEx {
	...
		CToolTipTextProviderList m_tooltipProvider;
	...
		virtual BOOL GetToolbarButtonToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
			return m_tooltipProvider.OnGetToolTipText( pButton, strTTText );
		}
	...

- Instantiate (or inherit) CMFCPropertySheetToolTipTextProvider in the property sheet, passing CToolTipTextProviderList instance to its constructor.
- Override CMFCPropertySheetToolTipTextProvider::FillToolTipText, update text of the CString& passed, return TRUE if the text is updated.

	class ContextPropSheet : public CMFCPropertySheet, public CMFCPropertySheetToolTipTextProvider {
	...
		ContextPropSheet( CToolTipTextProviderList * list )
			: CMFCPropertySheet( L"Context", this )
			, CMFCPropertySheetToolTipTextProvider( list, this ) {}
	...
		virtual BOOL FillToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText ) {
			strTTText = L"Decorated button text: ==> ";
			strTTText += pButton->m_strText;
			strTTText += L" <==";
			return TRUE;
		}
	...

- If needed, implement another logic of tooltip processing using CToolTipTextProvider with OnGetToolTipText override.

*/


class CToolTipTextProviderList;

/*
* Basic processor of tooltip text requests.
*/
class CToolTipTextProvider {

protected:

	// Parent list.
	CToolTipTextProviderList * m_list;

public:
	
	// Constructor. If list is not NULL, calls Subscribe( list ).
	CToolTipTextProvider( CToolTipTextProviderList * list );
	
	// Destructor. Calls Unsubscribe().
	virtual ~CToolTipTextProvider();

	// Subscribe this instance to "need tooltip text" requests routed via CToolTipTextProviderList.
	// Separate method for cases when the list can't be provided in constructor.
	void Subscribe( CToolTipTextProviderList * list );

	// Unsubscribe this instance from "need tooltip text" requests.
	// Separate method for cases when the list must be unbound before destructor.
	void Unsubscribe();

	// This method is called for all "need tooltip text" requests, including those not related to this instance.
	// The method should filter requests and process only those, which are related to this instance.
	// Should update strTTText and return TRUE if the request for the tooltip text is processed.
	virtual BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );

	virtual BOOL OnGetMessageString( UINT nID, CString & rMessage );
};


/*
* A CList-based list of tooltip text requests processors.
*/
class CToolTipTextProviderList: public CList < CToolTipTextProvider *, CToolTipTextProvider * > {

public:

	CToolTipTextProviderList();
	virtual ~CToolTipTextProviderList();

	// Adds item to the list.
	void AddProvider( CToolTipTextProvider * item );

	// Removes item from the list.
	void RemoveProvider( CToolTipTextProvider * item );

	// Should be called from CMDIFrameWndEx::GetToolbarButtonToolTipText.
	// Translates the tooltip text request to all registered processors.
	BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );

	void OnGetMessageString( UINT nID, CString & rMessage ) const;
};


/*
* CMFCPropertySheet-specific processor of tooltip text requests.
*/
class CMFCPropertySheetToolTipTextProvider : public CToolTipTextProvider {

protected:

	// Parent sheet.
	CMFCPropertySheet * m_pPropertySheet;

	CMap < UINT, UINT, CMFCToolBarButton *, CMFCToolBarButton * > m_buttonsByID;

	BOOL PrepareContext( CMFCToolBarButton * pToolBarButton, CMFCOutlookBarPaneButton *& pOutlookBarPaneButton, CMFCPropertyPage *& pPropertyPage );

	// Contains CMFCPropertySheet-specific logic. Calls FillToolTipText if notification is related to m_pPropertySheet.
	virtual BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );

	virtual BOOL OnGetMessageString( UINT nID, CString & rMessage );

public:

	// Subscribes to "need tooltip text" requests routed via CToolTipTextProviderList.
	// Saves CMFCPropertySheet ponter to filter "need tooltip text" requests.
	CMFCPropertySheetToolTipTextProvider( CToolTipTextProviderList * list, CMFCPropertySheet * propertySheet );

	// This method is called for all "need tooltip text" requests related to the CMFCPropertySheet passed via constructor.
	// Should update strTTText and return TRUE if the request for the tooltip text is processed.
	// The related button and property page objects are deduced and are passed to the function.
	// Note that HWND-related structures in pPropertyPage can be in non-initialized state if the corresponding page was not activated yet.
	virtual BOOL FillToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText );

	virtual BOOL FillToolTipDescription( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & rMessage );
};
