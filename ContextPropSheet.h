#pragma once
#include <afxpropertysheet.h>

class MainFrame;

class ContextPropSheet :
    public CMFCPropertySheet
{
protected:
  MainFrame * m_mainFrame;
public:
  ContextPropSheet( MainFrame * mainFrame );
  virtual ~ContextPropSheet();
  virtual BOOL OnInitDialog();
  BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );

  DECLARE_MESSAGE_MAP()
};

