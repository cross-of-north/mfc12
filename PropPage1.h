#pragma once
#include "afxdialogex.h"
#include "MyListCtrl.h"


// PropPage1 dialog

class PropPage1 : public CMFCPropertyPage
{
	DECLARE_DYNAMIC(PropPage1)

public:
	PropPage1(const int iResourceID, const CString stringID, CWnd* pParent = nullptr);   // standard constructor
	virtual ~PropPage1();

	CString m_stringID;
	
// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PROPPAGE_LARGE };
#endif

	CImageList ctrllistItems;
	MyListCtrl  myListCtrl { ctrllistItems };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
};
