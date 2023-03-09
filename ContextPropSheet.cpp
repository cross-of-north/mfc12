#include "pch.h"
#include "Resource.h"
#include "ContextPropSheet.h"

BEGIN_MESSAGE_MAP( ContextPropSheet, CMFCPropertySheet )
END_MESSAGE_MAP()

ContextPropSheet::ContextPropSheet()
  : CMFCPropertySheet( L"Context", this )
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
