#pragma once
#include <afxpropertysheet.h>
#include "ToolTipTextProvider.h"

class MainFrame;

class ContextPropSheet :
    public CMFCPropertySheet,
    public CMFCPropertySheetToolTipTextProvider
{
public:
  ContextPropSheet( CToolTipTextProviderList * list );
  virtual BOOL OnInitDialog();
  virtual BOOL FillToolTipText( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CString & strTTText );

  DECLARE_MESSAGE_MAP()
};

