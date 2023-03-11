#pragma once
#include <afxtempl.h>

/*
* Classes to customize tooltip appearance when such customization is not explicitly supported by MFC core.
*/

/*

1. Tooltip main text.
The problem is that the custom tooltip text messages are sent not to the controls the mouse is howering over,
but to the top-level main frame of the application.
To customize tooltip texts at the related control level, we do the following.
We are hooking the "get tooltip text" message at the main frame level.
There is a function for this already, CMDIFrameWndEx::GetToolbarButtonToolTipText.
Then we route the message to the related control (provided the control is subcribed for this message processing).

2. Tooltip description.
The description (additional text at the toolbar) is obtained in a similar way, via main frame of the application.
We are hooking the "get tooltip description text" message at the main frame level.
The hooked function is CFrameWnd::GetMessageString, it is designed to return description for toolbar buttons,
and instead of button/component pointer it receives command ID.
Buttons of CMFCOutlookBarPane have no command IDs internally, so we are generating command IDs for them, as
the part of CMDIFrameWndEx::GetToolbarButtonToolTipText processing activity.

3. Tooltip icon.
MFC draws icons in the tooltip only when the tooltip is bound to the CMFCRibbonButton or when the undelying toolbar
have the CImageList associated. CMFCOutlookBarPane have no CMFCRibbonButtons and have no image list in an invisible toolbar.
To show the icon in the tooltip we create an invisible CMFCRibbonButton containing the icon needed.

Usage:

- Instantiate CToolTipDataProviderList in the main frame context.

- Override CMDIFrameWndEx::GetToolbarButtonToolTipText at the main frame class,
  translate the call to the CToolTipDataProviderList::OnGetToolTipText.

- Override CFrameWnd::GetMessageString at the main frame class,
  translate the call to the CToolTipDataProviderList::GetMessageString.

	class MainFrame : public CMDIFrameWndEx {
	...
		CToolTipDataProviderList m_tooltipProvider;
	...
		virtual BOOL GetToolbarButtonToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
			return m_tooltipProvider.OnGetToolTipText( pButton, strTTText );
		}
	...
		virtual void GetMessageString( UINT nID, CString & rMessage ) const {
			m_tooltipProvider.OnGetMessageString( nID, rMessage );
		}
	...

- Instantiate (or inherit) CMFCPropertySheetToolTipDataProvider in the property sheet,
  passing CToolTipDataProviderList instance to its constructor.

- Override CMFCPropertySheetToolTipDataProvider::FillToolTipProperties, update CToolTipDataProviderProperties& passed,
  (tooltip texts, icon handle), return TRUE if properties should be applied.

	class ContextPropSheet : public CMFCPropertySheet, public CMFCPropertySheetToolTipDataProvider {
	...
		ContextPropSheet( CToolTipDataProviderList * list )
			: CMFCPropertySheet( L"Context", this )
			, CMFCPropertySheetToolTipDataProvider( list, this ) {}
	...
		virtual BOOL FillToolTipProperties( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CToolTipDataProviderProperties & props ) {
			props.m_strText = L"Decorated button text: ==> ";
			props.m_strText += pButton->m_strText;
			props.m_strText += L" <==";
			return TRUE;
		}
	...

- If needed, implement another logic of tooltip processing using CToolTipDataProvider with FillToolTipProperties override.

*/


// forward declarations
class CToolTipDataProviderList;
class CCustomMFCRibbonButton;

/*
* Tooltip parameters container.
*/
class CToolTipDataProviderProperties {
public:
	CToolTipDataProviderProperties( CMFCToolBarButton * pButton );
	~CToolTipDataProviderProperties();

	// text
	CString m_strText;

	// description
	CString m_strDescription;

	// icon
	HICON m_hIcon;

	// MFC tooltip properties
	CMFCToolTipInfo m_props;
	
	// fields for the internal usage

	// tooltip control to apply MFC properties to (can't be deduced from callbacks)
	CMFCToolTipCtrl * m_pToolTip;
	
	// pointer to the button (to map command ID to)
	CMFCToolBarButton * m_pButton;
	
	// CMFCRibbonButton owning pointer
	CCustomMFCRibbonButton * m_pRibbonButton;

private:
	// declared but not defined
	CToolTipDataProviderProperties( const CToolTipDataProviderProperties & );
	CToolTipDataProviderProperties & operator=( const CToolTipDataProviderProperties & );
};

/*
* Basic processor of tooltip data requests.
*/
class CToolTipDataProvider {

protected:

	// Parent list.
	CToolTipDataProviderList * m_list;

	// Tooltip properties storage.
	// Owns CToolTipDataProviderProperties instances.
	CMap < UINT, UINT, CToolTipDataProviderProperties *, CToolTipDataProviderProperties * > m_toolTipDataMap;

public:
	
	// Constructor. If list is not NULL, calls Subscribe( list ).
	CToolTipDataProvider( CToolTipDataProviderList * list );
	
	// Destructor. Calls Unsubscribe().
	virtual ~CToolTipDataProvider();

	// Subscribe this instance to "need tooltip data" requests routed via CToolTipDataProviderList.
	// Separate method for cases when the list can't be provided in constructor.
	void Subscribe( CToolTipDataProviderList * list );

	// Unsubscribe this instance from "need tooltip data" requests.
	// Separate method for cases when the list must be unbound before destructor.
	void Unsubscribe();

	// This method is called for all "need tooltip text" requests, including those not related to this instance.
	// The method filters requests and process only those, which are related to this instance.
	// Should update strTTText and return TRUE if the request is processed.
	// The default implementation calls FillToolTipProperties and does tooltip initialization.
	virtual BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );

	// This method is called for all "need tooltip description" requests, including those not related to this instance.
	// The method filters requests and process only those, which are related to this instance.
	// Should update rMessage and return TRUE if the request is processed.
	// The default implementation uses text from m_toolTipDataMap pre-filled in OnGetToolTipText().
	virtual BOOL OnGetMessageString( UINT nID, CString & rMessage );

	// The method is called to fill tooltip properties.
	// Should update props and return TRUE if properties should be applied.
	// The default implementation does nothing.
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

	// Should be called from CFrameWnd::GetMessageString.
	// Translates the tooltip description text request to all registered processors.
	void OnGetMessageString( UINT nID, CString & rMessage ) const;
};


/*
* CMFCPropertySheet-specific processor of tooltip data requests.
*/
class CMFCPropertySheetToolTipDataProvider : public CToolTipDataProvider {

protected:

	// Parent sheet.
	CMFCPropertySheet * m_pPropertySheet;

	// Locates parent MFC structures of button provided that button is a child of this property sheet.
	BOOL PrepareContext( CMFCToolBarButton * pToolBarButton, CMFCOutlookBarPaneButton *& pOutlookBarPaneButton, CMFCPropertyPage *& pPropertyPage );

	// Contains CMFCPropertySheet-specific logic.
	// Fills props.m_pToolTip field with a pointer to the the corresponding object owned by m_pPropertySheet.
	// Calls public overloaded FillToolTipProperties() variation if notification is related to m_pPropertySheet.
	virtual BOOL FillToolTipProperties( CToolTipDataProviderProperties & props );

public:

	// Subscribes to "need tooltip data" requests routed via CToolTipDataProviderList.
	// Stores CMFCPropertySheet ponter to filter "need tooltip data" requests.
	CMFCPropertySheetToolTipDataProvider( CToolTipDataProviderList * list, CMFCPropertySheet * propertySheet );

	// This method is called for all "need tooltip data" requests related to the CMFCPropertySheet passed via constructor.
	// Should update props and return TRUE if the request for the tooltip text is processed.
	// Note that HWND-related structures in pPropertyPage can be in non-initialized state if the corresponding page was not activated yet.
	virtual BOOL FillToolTipProperties( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CToolTipDataProviderProperties & props );

};
