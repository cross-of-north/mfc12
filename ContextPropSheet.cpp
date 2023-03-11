#include "pch.h"
#include "Resource.h"
#include "ContextPropSheet.h"

#include "MainFrm.h"
#include "PropPage1.h"

BEGIN_MESSAGE_MAP( ContextPropSheet, CMFCPropertySheet )
END_MESSAGE_MAP()

ContextPropSheet::ContextPropSheet( CToolTipTextProviderList * list )
	: CMFCPropertySheet( L"Context", this )
	, CMFCPropertySheetToolTipTextProvider( list, this )
{
	SetLook( CMFCPropertySheet::PropSheetLook_OutlookBar ); // outlookbar mode not tree or list so page icons will show up and so tooltips will appear
	SetIconsList( IDB_CONTEXT_LARGE, 32 );
	EnableToolTips();
}

BOOL ContextPropSheet::OnInitDialog()
{
	BOOL oKay { TRUE };
	BOOL bResult = CMFCPropertySheet::OnInitDialog();

	return bResult;
}

// the hack to access internal structures from here
class friendly_CMFCOutlookBarPaneList: public CMFCOutlookBarPaneList {
	friend class ContextPropSheet;
};
class friendly_CMFCToolTipCtrl: public CMFCToolTipCtrl {
	friend class ContextPropSheet;
};

class CCustomMFCRibbonButton: public CMFCRibbonButton {
public:
	CCustomMFCRibbonButton( UINT nID, LPCTSTR lpszText, HICON hIcon, BOOL bAlwaysShowDescription = FALSE, HICON hIconSmall = NULL, BOOL bAutoDestroyIcon = FALSE, BOOL bAlphaBlendIcon = FALSE )
		: CMFCRibbonButton( nID, lpszText, hIcon, bAlwaysShowDescription, hIconSmall, bAutoDestroyIcon, bAlphaBlendIcon )
	{
		m_bIsLargeImage = TRUE;
	}
};

BOOL ContextPropSheet::FillToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText ) {
	CToolTipCtrl * pToolTip = ( static_cast < friendly_CMFCOutlookBarPaneList * > ( &m_wndPane1 ) )->m_pToolTip;
	CMFCToolTipCtrl * _pMFCToolTip = dynamic_cast < CMFCToolTipCtrl * > ( pToolTip );
	friendly_CMFCToolTipCtrl * pMFCToolTip = static_cast < friendly_CMFCToolTipCtrl * > ( _pMFCToolTip );
	int iPageIndex = GetPageIndex( pPropertyPage );
	if ( iPageIndex != -1 ) {
		HICON hIcon = m_Icons.ExtractIconW( iPageIndex );
		CMFCRibbonButton * pRibbonButton = new CCustomMFCRibbonButton( 0, L"", hIcon ); // leak
		pMFCToolTip->m_pRibbonButton = pRibbonButton;
	}
	PropPage1 * pp1 = dynamic_cast < PropPage1 * > ( pPropertyPage );
	if ( pp1 != NULL ) {
		strTTText = pp1->m_stringID;
	}
	return TRUE;
}

BOOL ContextPropSheet::FillToolTipDescription( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & rMessage ) {
	rMessage = m_strCaption;
	rMessage += ", ";
	rMessage += pButton->m_strText;
	rMessage += " tooltip at time ";
	rMessage += CTime::GetCurrentTime().Format( "%H:%M:%S" );
	return TRUE;
}