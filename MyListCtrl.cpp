// MyListCtrl.cpp : implementation file
//

#include "pch.h"
#include "MFC12.h"
#include "afxdialogex.h"
#include "MyListCtrl.h"


// MyListCtrl dialog

IMPLEMENT_DYNAMIC(MyListCtrl, CMFCListCtrl)

MyListCtrl::MyListCtrl( CImageList& ctrlListItems )
	: ctrlListItems { ctrlListItems }
{

}

MyListCtrl::~MyListCtrl()
{
}

void MyListCtrl::DoDataExchange(CDataExchange* pDX)
{
	CMFCListCtrl::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(MyListCtrl, CMFCListCtrl)
END_MESSAGE_MAP()


// MyListCtrl message handlers
