// MyListCtrl.cpp : implementation file
//

#include "pch.h"
#include "ToolTipTextProvider.h"


// CToolTipTextProvider

CToolTipTextProvider::CToolTipTextProvider( CToolTipTextProviderList * list )
: m_list( list )
{
    Subscribe( m_list );
}

CToolTipTextProvider::~CToolTipTextProvider() {
    Unsubscribe();
}

void CToolTipTextProvider::Subscribe( CToolTipTextProviderList * list ) {
    if ( list != NULL ) {
        m_list = list;
        m_list->AddProvider( this );
    }
}

void CToolTipTextProvider::Unsubscribe() {
    if ( m_list != NULL ) {
        m_list->RemoveProvider( this );
    }
}

BOOL CToolTipTextProvider::OnGetToolTipText( CMFCToolBarButton * /*pButton*/, CString & /*strTTText*/ ) {
    return FALSE;
}


// CToolTipTextProviderList

CToolTipTextProviderList::CToolTipTextProviderList() {
}

CToolTipTextProviderList::~CToolTipTextProviderList() {
}

BOOL CToolTipTextProviderList::OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
    BOOL bResult = FALSE;
    // calling items' OnGetToolTipText until TRUE is returned
    POSITION pos = GetHeadPosition();
    while ( pos && bResult == FALSE ) {
        bResult = GetNext( pos )->OnGetToolTipText( pButton, strTTText );
    }
    return bResult;
}

void CToolTipTextProviderList::AddProvider( CToolTipTextProvider * item ) {
    if ( Find( item ) == NULL ) {
        AddTail( item );
    }
}

void CToolTipTextProviderList::RemoveProvider( CToolTipTextProvider * item ) {
    POSITION pos = Find( item );
    if ( pos != NULL ) {
        RemoveAt( pos );
    }
}


// CMFCPropertySheetToolTipTextProvider

CMFCPropertySheetToolTipTextProvider::CMFCPropertySheetToolTipTextProvider( CToolTipTextProviderList * list, CMFCPropertySheet * pPropertySheet )
    : CToolTipTextProvider( list )
    , m_pPropertySheet( pPropertySheet )
{
}

// the hack to access CMFCPropertySheet internal structures from here
class friendly_CMFCPropertySheet: public CMFCPropertySheet {
    friend class CMFCPropertySheetToolTipTextProvider;
};

BOOL CMFCPropertySheetToolTipTextProvider::OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
    if ( pButton != NULL ) {
        
        // searching the parent sheet up to the windows hierarchy
        
        CWnd * pWnd = pButton->GetParentWnd();
        while ( pWnd != NULL ) {
            if ( pWnd == static_cast < CWnd * > ( m_pPropertySheet ) ) {
                
                // found
                
                // determining property page by the button index
                CMFCOutlookBarPaneList & pOutlookBarPaneList = static_cast < friendly_CMFCPropertySheet * > ( m_pPropertySheet )->m_wndPane1;
                int buttonIndex = pOutlookBarPaneList.ButtonToIndex( pButton );
                if ( buttonIndex != -1 ) {
                    CMFCPropertyPage * pPropertyPage = dynamic_cast < CMFCPropertyPage * > ( m_pPropertySheet->GetPage( buttonIndex ) );
                    CMFCOutlookBarPaneButton * pOutlookButton = dynamic_cast < CMFCOutlookBarPaneButton * > ( pButton );
                    if ( 
                        pPropertyPage != NULL && 
                        pOutlookButton != NULL && 
                        // Now that we know everything related to the tooltip needed, we are obtaining the tooltip text.
                        FillToolTipText( pOutlookButton, pPropertyPage, strTTText ) == TRUE
                     ) {
                        return TRUE;
                    }
                }

            }
            pWnd = pWnd->GetParent();
        };
    }
    return FALSE;
}

BOOL CMFCPropertySheetToolTipTextProvider::FillToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText ) {
    return FALSE;
}