#include "pch.h"
#include "Resource.h"
#include "ContextPropSheet.h"

#include "MainFrm.h"

BEGIN_MESSAGE_MAP( ContextPropSheet, CMFCPropertySheet )
END_MESSAGE_MAP()

ContextPropSheet::ContextPropSheet( MainFrame * mainFrame )
  : CMFCPropertySheet( L"Context", this )
  , m_mainFrame( mainFrame )
{
  SetLook( CMFCPropertySheet::PropSheetLook_OutlookBar ); // outlookbar mode not tree or list so page icons will show up and so tooltips will appear
  SetIconsList( IDB_CONTEXT_LARGE, 32 );
  EnableToolTips();
}

ContextPropSheet::~ContextPropSheet() {
  m_mainFrame->UnsubscribeToolTipProvider( this );
}

BOOL ContextPropSheet::OnInitDialog()
{
  BOOL oKay { TRUE };
  BOOL bResult = CMFCPropertySheet::OnInitDialog();
  m_mainFrame->SubscribeToolTipProvider( this );
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
        if ( pOutlookButton != NULL ) {
          strTTText = pButton->m_strText;
          strTTText += " tooltip at time ";
          strTTText += CTime::GetCurrentTime().Format( "%H:%M:%S" );
          return TRUE;
        }
      }
      pWnd = pWnd->GetParent();
    };
  }
  return FALSE;
}
