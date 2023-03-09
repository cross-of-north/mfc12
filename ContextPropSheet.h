#pragma once
#include <afxpropertysheet.h>
class ContextPropSheet :
    public CMFCPropertySheet
{
public:
  ContextPropSheet();
  virtual BOOL OnInitDialog();

  DECLARE_MESSAGE_MAP()
};

