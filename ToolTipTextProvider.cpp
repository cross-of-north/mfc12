
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

BOOL CToolTipTextProvider::OnGetMessageString( UINT /*nID*/, CString & /*rMessage*/ ) {
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

void CToolTipTextProviderList::OnGetMessageString( UINT nID, CString & rMessage ) const {
    // calling items' OnGetMessageString until TRUE is returned
    POSITION pos = GetHeadPosition();
    while ( pos && GetNext( pos )->OnGetMessageString( nID, rMessage ) == FALSE ) {
        // continue
    }
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

BOOL CMFCPropertySheetToolTipTextProvider::PrepareContext( CMFCToolBarButton * pToolBarButton, CMFCOutlookBarPaneButton * &pOutlookBarPaneButton, CMFCPropertyPage * &pPropertyPage ) {
    pOutlookBarPaneButton = NULL;
    pPropertyPage = NULL;

    if ( pToolBarButton != NULL ) {

        // searching the parent sheet up to the windows hierarchy

        CWnd * pWnd = pToolBarButton->GetParentWnd();
        while ( pWnd != NULL ) {
            if ( pWnd == static_cast < CWnd * > ( m_pPropertySheet ) ) {

                // found

                // determining property page by the button index
                CMFCOutlookBarPaneList & pOutlookBarPaneList = static_cast < friendly_CMFCPropertySheet * > ( m_pPropertySheet )->m_wndPane1;
                int buttonIndex = pOutlookBarPaneList.ButtonToIndex( pToolBarButton );
                if ( buttonIndex != -1 ) {
                    pPropertyPage = dynamic_cast < CMFCPropertyPage * > ( m_pPropertySheet->GetPage( buttonIndex ) );
                    pOutlookBarPaneButton = dynamic_cast < CMFCOutlookBarPaneButton * > ( pToolBarButton );
                    break;
                }

            }
            pWnd = pWnd->GetParent();
        };
    }

    return pOutlookBarPaneButton != NULL && pPropertyPage != NULL;
}


BOOL CMFCPropertySheetToolTipTextProvider::OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
    CMFCPropertyPage * pPropertyPage = NULL;
    CMFCOutlookBarPaneButton * pOutlookButton = NULL;
    if ( PrepareContext( pButton, pOutlookButton, pPropertyPage ) ) {
        // Now that we know everything related to the tooltip needed, we are obtaining the tooltip text.
        if ( pButton->m_nID == 0 ) {
            pButton->m_nID = PtrToUint( pButton ); // unique low-DWORD-adress-based command
            m_buttonsByID[ pButton->m_nID ] = pButton;
        }
        return FillToolTipText( pOutlookButton, pPropertyPage, strTTText );
    }
    return FALSE;
}

BOOL CMFCPropertySheetToolTipTextProvider::OnGetMessageString( UINT nID, CString & rMessage ) {
    CMFCToolBarButton * pButton = NULL;
    m_buttonsByID.Lookup( nID, pButton );
    if ( pButton != NULL ) {
        CMFCPropertyPage * pPropertyPage = NULL;
        CMFCOutlookBarPaneButton * pOutlookButton = NULL;
        if ( PrepareContext( pButton, pOutlookButton, pPropertyPage ) ) {
            return FillToolTipDescription( pOutlookButton, pPropertyPage, rMessage );
        }
    }
    return FALSE;
}

BOOL CMFCPropertySheetToolTipTextProvider::FillToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText ) {
    return FALSE;
}

BOOL CMFCPropertySheetToolTipTextProvider::FillToolTipDescription( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & rMessage ) {
    return FALSE;
}