
#include "pch.h"
#include "ToolTipDataProvider.h"

// hacks to access internal MFC structures from here
class friendly_CMFCToolTipCtrl: public CMFCToolTipCtrl {
    friend class CToolTipDataProvider;
};
class friendly_CMFCOutlookBarPaneList: public CMFCOutlookBarPaneList {
    friend class CMFCPropertySheetToolTipDataProvider;
};
class friendly_CMFCPropertySheet: public CMFCPropertySheet {
    friend class CMFCPropertySheetToolTipDataProvider;
};

// This class created only to set m_bIsLargeImage = TRUE.
// Since the instances of CMFCRibbonButton are created here, it is possible to do without "friendly" hacks.
class CCustomMFCRibbonButton: public CMFCRibbonButton {
public:
    CCustomMFCRibbonButton( UINT nID, LPCTSTR lpszText, HICON hIcon, BOOL bAlwaysShowDescription = FALSE, HICON hIconSmall = NULL, BOOL bAutoDestroyIcon = FALSE, BOOL bAlphaBlendIcon = FALSE )
        : CMFCRibbonButton( nID, lpszText, hIcon, bAlwaysShowDescription, hIconSmall, bAutoDestroyIcon, bAlphaBlendIcon ) {
        m_bIsLargeImage = TRUE;
    }
    virtual ~CCustomMFCRibbonButton() {
    }
};


// CToolTipDataProviderProperties

CToolTipDataProviderProperties::CToolTipDataProviderProperties( CMFCToolBarButton * pButton )
    : m_hIcon( NULL )
    , m_pButton( pButton )
    , m_pToolTip( NULL )
    , m_pRibbonButton( NULL )
{
}

CToolTipDataProviderProperties::~CToolTipDataProviderProperties() {
    delete m_pRibbonButton;
}


// CToolTipDataProvider

CToolTipDataProvider::CToolTipDataProvider( CToolTipDataProviderList * list )
    : m_list( list )
{
    Subscribe( m_list );
}

CToolTipDataProvider::~CToolTipDataProvider() {
    Unsubscribe();

    // delete tooltip props structures
    POSITION pos = m_toolTipDataMap.GetStartPosition();
    while ( pos != NULL ) {
        UINT key = 0;
        CToolTipDataProviderProperties * props = NULL;
        m_toolTipDataMap.GetNextAssoc( pos, key, props );
        delete props;
    }
}

void CToolTipDataProvider::Subscribe( CToolTipDataProviderList * list ) {
    if ( list != NULL ) {
        m_list = list;
        m_list->AddProvider( this );
    }
}

void CToolTipDataProvider::Unsubscribe() {
    if ( m_list != NULL ) {
        m_list->RemoveProvider( this );
    }
}

BOOL CToolTipDataProvider::OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {

    // This is the first chance to perform tooltip-related activity.
    // We are creating tooltip properies storage and call FillToolTipProperties to get customized data.
    CToolTipDataProviderProperties * props = new CToolTipDataProviderProperties( pButton );
    BOOL bResult = FillToolTipProperties( *props );

    if ( bResult ) {

        // If properties data is customized, we save it for later usage.

        // Not all buttons have command IDs set, but we need ID to differentiate buttons in OnGetMessageString().
        // When there is no command ID, we set low DWORD of button address (sufficiently unique) as a command ID.
        if ( pButton->m_nID == 0 ) {
            pButton->m_nID = PtrToUint( pButton );
        }

        // Overwriting old props value.
        CToolTipDataProviderProperties * old_props = NULL;
        m_toolTipDataMap.Lookup( pButton->m_nID, old_props );
        if ( old_props != NULL ) {
            delete old_props;
        }
        m_toolTipDataMap[ pButton->m_nID ] = props;

        // Setting toolbar text from props.
        strTTText = props->m_strText;

        if ( props->m_pToolTip != NULL ) {
            
            // Updating icon.
            friendly_CMFCToolTipCtrl * pToolTip = static_cast < friendly_CMFCToolTipCtrl * > ( props->m_pToolTip );
            
            if ( props->m_hIcon != NULL ) {
                // Creating CMFCRibbonButton instance is the only way to provide the tooltip icon here.
                props->m_pRibbonButton = new CCustomMFCRibbonButton( 0, L"", props->m_hIcon );
                pToolTip->m_pRibbonButton = props->m_pRibbonButton;
            }

            // Updating other parameters.
            props->m_pToolTip->SetParams( &props->m_props );
        }

    } else {

        // No need to store data if the tooltip is not customized.
        delete props;

    }

    return bResult;
}

BOOL CToolTipDataProvider::OnGetMessageString( UINT nID, CString & rMessage ) {
    
    // Looking up tooltip properies by command ID
    CToolTipDataProviderProperties * props = NULL;
    m_toolTipDataMap.Lookup( nID, props );

    if ( props != NULL && !props->m_strDescription.IsEmpty() ) {
        // If button is found, using description saved while processing OnGetToolTipText (which always comes before OnGetMessageString).
        rMessage = props->m_strDescription;
        return TRUE;
    }
    return FALSE;
}

BOOL CToolTipDataProvider::FillToolTipProperties( CToolTipDataProviderProperties & props ) {
    return FALSE;
}


// CToolTipDataProviderList

CToolTipDataProviderList::CToolTipDataProviderList() {
}

CToolTipDataProviderList::~CToolTipDataProviderList() {
}

BOOL CToolTipDataProviderList::OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
    BOOL bResult = FALSE;
    // calling items' OnGetToolTipText until TRUE is returned
    POSITION pos = GetHeadPosition();
    while ( pos != NULL && bResult == FALSE ) {
        bResult = GetNext( pos )->OnGetToolTipText( pButton, strTTText );
    }
    return bResult;
}

void CToolTipDataProviderList::OnGetMessageString( UINT nID, CString & rMessage ) const {
    // calling items' OnGetMessageString until TRUE is returned
    POSITION pos = GetHeadPosition();
    while ( pos != NULL && GetNext( pos )->OnGetMessageString( nID, rMessage ) == FALSE ) {
        // continue
    }
}

void CToolTipDataProviderList::AddProvider( CToolTipDataProvider * item ) {
    if ( Find( item ) == NULL ) {
        AddTail( item );
    }
}

void CToolTipDataProviderList::RemoveProvider( CToolTipDataProvider * item ) {
    POSITION pos = Find( item );
    if ( pos != NULL ) {
        RemoveAt( pos );
    }
}


// CMFCPropertySheetToolTipDataProvider

CMFCPropertySheetToolTipDataProvider::CMFCPropertySheetToolTipDataProvider( CToolTipDataProviderList * list, CMFCPropertySheet * pPropertySheet )
    : CToolTipDataProvider( list )
    , m_pPropertySheet( pPropertySheet )
{
}

BOOL CMFCPropertySheetToolTipDataProvider::PrepareContext( CMFCToolBarButton * pToolBarButton, CMFCOutlookBarPaneButton * &pOutlookBarPaneButton, CMFCPropertyPage * &pPropertyPage ) {
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

BOOL CMFCPropertySheetToolTipDataProvider::FillToolTipProperties( CToolTipDataProviderProperties & props ) {
    CMFCPropertyPage * pPropertyPage = NULL;
    CMFCOutlookBarPaneButton * pOutlookButton = NULL;
    if ( PrepareContext( props.m_pButton, pOutlookButton, pPropertyPage ) ) {
        
        // Storing tooltip class instance in parameters.
        friendly_CMFCPropertySheet * pPropertySheet = static_cast < friendly_CMFCPropertySheet * > ( m_pPropertySheet );
        CToolTipCtrl * pToolTip = ( static_cast < friendly_CMFCOutlookBarPaneList * > ( &pPropertySheet->m_wndPane1 ) )->m_pToolTip;
        props.m_pToolTip = dynamic_cast < CMFCToolTipCtrl * > ( pToolTip );
        
        // Now we know everything related to this tooltip.
        // Updating tooltip properties.
        return FillToolTipProperties( pOutlookButton, pPropertyPage, props );
    }
    return FALSE;
}

BOOL CMFCPropertySheetToolTipDataProvider::FillToolTipProperties( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CToolTipDataProviderProperties & props ) {
    return FALSE;
}
