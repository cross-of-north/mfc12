// MyListCtrl.cpp : implementation file
//

#include "pch.h"
#include "ToolTipTextProvider.h"


// CToolTipTextProviderItem

CToolTipTextProviderItem::CToolTipTextProviderItem( CToolTipTextProviderList * list )
: m_list( list )
{
    Subscribe( m_list );
}

CToolTipTextProviderItem::~CToolTipTextProviderItem() {
    Unsubscribe();
}

void CToolTipTextProviderItem::Subscribe( CToolTipTextProviderList * list ) {
    if ( list != NULL ) {
        m_list = list;
        m_list->AddProvider( this );
    }
}

void CToolTipTextProviderItem::Unsubscribe() {
    if ( m_list != NULL ) {
        m_list->RemoveProvider( this );
    }
}

BOOL CToolTipTextProviderItem::OnGetToolTipText( CMFCToolBarButton * /*pButton*/, CString & /*strTTText*/ ) {
    return FALSE;
}


// CToolTipTextProviderList

CToolTipTextProviderList::CToolTipTextProviderList() {
}

CToolTipTextProviderList::~CToolTipTextProviderList() {
}

BOOL CToolTipTextProviderList::OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
    BOOL bResult = FALSE;
    POSITION pos = GetHeadPosition();
    while ( pos && bResult == FALSE ) {
        bResult = GetNext( pos )->OnGetToolTipText( pButton, strTTText );
    }
    return bResult;
}

void CToolTipTextProviderList::AddProvider( CToolTipTextProviderItem * item ) {
    if ( Find( item ) == NULL ) {
        AddTail( item );
    }
}

void CToolTipTextProviderList::RemoveProvider( CToolTipTextProviderItem * item ) {
    POSITION pos = Find( item );
    if ( pos != NULL ) {
        RemoveAt( pos );
    }
}
