#pragma once
#include <afxpropertysheet.h>
#include "ToolTipTextProvider.h"

class MainFrame;

class ContextPropSheet :
    public CMFCPropertySheet,
    public CToolTipTextProviderItem
{
public:
  ContextPropSheet( CToolTipTextProviderList * list );
  virtual BOOL OnInitDialog();
  virtual BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );
  virtual BOOL FillToolTipText( CMFCOutlookBarPaneButton * pButton, CString & strTTText );

  DECLARE_MESSAGE_MAP()
};

