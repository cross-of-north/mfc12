#include "pch.h"
#include "Resource.h"
#include "ContextPropSheet.h"

#include "MainFrm.h"
#include "PropPage1.h"

BEGIN_MESSAGE_MAP( ContextPropSheet, CMFCPropertySheet )
END_MESSAGE_MAP()

ContextPropSheet::ContextPropSheet( CToolTipDataProviderList * list )
	: CMFCPropertySheet( L"Context", this )
	, CMFCPropertySheetToolTipDataProvider( list, this )
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

BOOL ContextPropSheet::FillToolTipProperties( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CToolTipDataProviderProperties & props ) {

	// tooltip text
	PropPage1 * pp1 = dynamic_cast < PropPage1 * > ( pPropertyPage );
	if ( pp1 != NULL ) {
		props.m_strText = pp1->m_stringID;
	}

	// tooltip description
	props.m_strDescription = m_strCaption;
	props.m_strDescription += ", ";
	props.m_strDescription += pButton->m_strText;
	props.m_strDescription += " tooltip at time ";
	props.m_strDescription += CTime::GetCurrentTime().Format( "%H:%M:%S" );

	// tooltip icon
	int iPageIndex = GetPageIndex( pPropertyPage );
	if ( iPageIndex != -1 ) {
		props.m_hIcon = m_Icons.ExtractIconW( iPageIndex );
	}

	// tooltip width limit
	props.m_props.m_nMaxDescrWidth = GetSystemMetrics( SM_CXSCREEN ) / 3;

	// do not draw separator
	props.m_props.m_bDrawSeparator = FALSE;

	return TRUE;
}
