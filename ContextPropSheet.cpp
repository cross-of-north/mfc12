#include "pch.h"
#include "Resource.h"
#include "ContextPropSheet.h"

#include "MainFrm.h"

BEGIN_MESSAGE_MAP( ContextPropSheet, CMFCPropertySheet )
END_MESSAGE_MAP()

ContextPropSheet::ContextPropSheet( CToolTipTextProviderList * list )
	: CMFCPropertySheet( L"Context", this )
	, CToolTipTextProviderItem( list ) {
	SetLook( CMFCPropertySheet::PropSheetLook_OutlookBar ); // outlookbar mode not tree or list so page icons will show up and so tooltips will appear
	SetIconsList( IDB_CONTEXT_LARGE, 32 );
	EnableToolTips();
}

BOOL ContextPropSheet::OnInitDialog() {
	BOOL oKay{ TRUE };
	BOOL bResult = CMFCPropertySheet::OnInitDialog();
	if ( m_wndOutlookBar.GetUnderlyingWindow() != NULL ) {
		m_wndOutlookBar.GetUnderlyingWindow()->EnableCustomToolTips( TRUE );
	}
	return bResult;
}

BOOL ContextPropSheet::OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText ) {
	if ( pButton != NULL ) {
		CWnd * pWnd = pButton->GetParentWnd();
		while ( pWnd != NULL ) {
			if ( pWnd == static_cast < CWnd * > ( this ) ) {
				CMFCOutlookBarPaneButton * pOutlookButton = dynamic_cast < CMFCOutlookBarPaneButton * > ( pButton );
				if ( FillToolTipText( pOutlookButton, strTTText ) == TRUE ) {
					return TRUE;
				}
			}
			pWnd = pWnd->GetParent();
		};
	}
	return FALSE;
}

BOOL ContextPropSheet::FillToolTipText( CMFCOutlookBarPaneButton * pButton, CString & strTTText ) {
	strTTText = pButton->m_strText;
	strTTText += " tooltip at time ";
	strTTText += CTime::GetCurrentTime().Format( "%H:%M:%S" );
	return TRUE;
}