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

BOOL ContextPropSheet::FillToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText ) {
	strTTText = m_strCaption;
	strTTText += ", ";
	strTTText += pButton->m_strText;
	strTTText += " tooltip at time ";
	strTTText += CTime::GetCurrentTime().Format( "%H:%M:%S" );
	PropPage1 * pp1 = dynamic_cast < PropPage1 * > ( pPropertyPage );
	if ( pp1 != NULL ) {
		strTTText += " (";
		strTTText += pp1->m_stringID;
		strTTText += ")";
	}
	return TRUE;
}