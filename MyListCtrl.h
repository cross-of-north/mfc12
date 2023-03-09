#pragma once
#include "afxdialogex.h"


// MyListCtrl dialog

class MyListCtrl : public CMFCListCtrl
{
	DECLARE_DYNAMIC(MyListCtrl)

public:
	MyListCtrl( CImageList& ctrlListItems );   // standard constructor
	virtual ~MyListCtrl();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PROPPAGE_LARGE };
#endif

	CImageList& ctrlListItems;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};
