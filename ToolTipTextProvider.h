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

- Instantiate CToolTipDataProviderList in the main frame context.
- Override CMDIFrameWndEx::GetToolbarButtonToolTipText at the main frame class, translate the call to the CToolTipDataProviderList::OnGetToolTipText.

	class MainFrame : public CMDIFrameWndEx {
	...
		CToolTipDataProviderList m_tooltipProvider;
	...
		virtual BOOL GetToolbarButtonToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
			return m_tooltipProvider.OnGetToolTipText( pButton, strTTText );
		}
	...

- Instantiate (or inherit) CMFCPropertySheetToolTipDataProvider in the property sheet, passing CToolTipDataProviderList instance to its constructor.
- Override CMFCPropertySheetToolTipDataProvider::GetToolTipText, update text of the CString& passed, return TRUE if the text is updated.

	class ContextPropSheet : public CMFCPropertySheet, public CMFCPropertySheetToolTipDataProvider {
	...
		ContextPropSheet( CToolTipDataProviderList * list )
			: CMFCPropertySheet( L"Context", this )
			, CMFCPropertySheetToolTipDataProvider( list, this ) {}
	...
		virtual BOOL GetToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText ) {
			strTTText = L"Decorated button text: ==> ";
			strTTText += pButton->m_strText;
			strTTText += L" <==";
			return TRUE;
		}
	...

- If needed, implement another logic of tooltip processing using CToolTipDataProvider with OnGetToolTipText override.

*/


class CToolTipDataProviderList;

class CCustomMFCRibbonButton;

class CToolTipDataProviderProperties {
private:
	// declared but not defined
	CToolTipDataProviderProperties( const CToolTipDataProviderProperties & );
	CToolTipDataProviderProperties & operator=( const CToolTipDataProviderProperties & );
public:
	CToolTipDataProviderProperties( CMFCToolBarButton * pButton );
	~CToolTipDataProviderProperties();

	CString m_strText;
	CString m_strDescription;
	HICON m_hIcon;
	CMFCToolTipInfo m_props;
	CMFCToolBarButton * m_pButton;
	CMFCToolTipCtrl * m_pToolTip;
	CCustomMFCRibbonButton * m_pRibbonButton;
};

/*
* Basic processor of tooltip text requests.
*/
class CToolTipDataProvider {

protected:

	// Parent list.
	CToolTipDataProviderList * m_list;

	CMap < UINT, UINT, CToolTipDataProviderProperties *, CToolTipDataProviderProperties * > m_toolTipDataMap;

public:
	
	// Constructor. If list is not NULL, calls Subscribe( list ).
	CToolTipDataProvider( CToolTipDataProviderList * list );
	
	// Destructor. Calls Unsubscribe().
	virtual ~CToolTipDataProvider();

	// Subscribe this instance to "need tooltip text" requests routed via CToolTipDataProviderList.
	// Separate method for cases when the list can't be provided in constructor.
	void Subscribe( CToolTipDataProviderList * list );

	// Unsubscribe this instance from "need tooltip text" requests.
	// Separate method for cases when the list must be unbound before destructor.
	void Unsubscribe();

	// This method is called for all "need tooltip text" requests, including those not related to this instance.
	// The method should filter requests and process only those, which are related to this instance.
	// Should update strTTText and return TRUE if the request for the tooltip text is processed.
	virtual BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );

	virtual BOOL OnGetMessageString( UINT nID, CString & rMessage );

	virtual BOOL FillToolTipProperties( CToolTipDataProviderProperties & props );
};


/*
* A CList-based list of tooltip text requests processors.
*/
class CToolTipDataProviderList: public CList < CToolTipDataProvider *, CToolTipDataProvider * > {

public:

	CToolTipDataProviderList();
	virtual ~CToolTipDataProviderList();

	// Adds item to the list.
	void AddProvider( CToolTipDataProvider * item );

	// Removes item from the list.
	void RemoveProvider( CToolTipDataProvider * item );

	// Should be called from CMDIFrameWndEx::GetToolbarButtonToolTipText.
	// Translates the tooltip text request to all registered processors.
	BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );

	void OnGetMessageString( UINT nID, CString & rMessage ) const;
};


/*
* CMFCPropertySheet-specific processor of tooltip text requests.
*/
class CMFCPropertySheetToolTipDataProvider : public CToolTipDataProvider {

protected:

	// Parent sheet.
	CMFCPropertySheet * m_pPropertySheet;

	BOOL PrepareContext( CMFCToolBarButton * pToolBarButton, CMFCOutlookBarPaneButton *& pOutlookBarPaneButton, CMFCPropertyPage *& pPropertyPage );

	// Contains CMFCPropertySheet-specific logic. Calls GetToolTipText if notification is related to m_pPropertySheet.
	virtual BOOL FillToolTipProperties( CToolTipDataProviderProperties & props );

public:

	// Subscribes to "need tooltip text" requests routed via CToolTipDataProviderList.
	// Saves CMFCPropertySheet ponter to filter "need tooltip text" requests.
	CMFCPropertySheetToolTipDataProvider( CToolTipDataProviderList * list, CMFCPropertySheet * propertySheet );

	// This method is called for all "need tooltip text" requests related to the CMFCPropertySheet passed via constructor.
	// Should update strTTText and return TRUE if the request for the tooltip text is processed.
	// The related button and property page objects are deduced and are passed to the function.
	// Note that HWND-related structures in pPropertyPage can be in non-initialized state if the corresponding page was not activated yet.
	virtual BOOL FillToolTipProperties( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CToolTipDataProviderProperties & props );

	/*
	virtual BOOL GetToolTipDescription( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & rMessage );

	virtual BOOL GetToolTipIcon( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, HICON & hIcon );

	virtual BOOL GetToolTipProperties( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CMFCToolTipInfo & props );
	*/
};
