#pragma once
#include <afxpropertysheet.h>
#include "ToolTipDataProvider.h"

class MainFrame;

class ContextPropSheet :
    public CMFCPropertySheet,
    public CMFCPropertySheetToolTipDataProvider
{
public:
  ContextPropSheet( CToolTipDataProviderList * list );
  virtual BOOL OnInitDialog();
  virtual BOOL FillToolTipProperties( CMFCOutlookBarPaneButton * pButton, CMFCPropertyPage * pPropertyPage, CToolTipDataProviderProperties & props );

  DECLARE_MESSAGE_MAP()
};

