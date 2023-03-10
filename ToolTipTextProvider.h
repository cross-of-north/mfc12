#pragma once
#include <afxtempl.h>

class CToolTipTextProviderList;

class CToolTipTextProviderItem {
protected:
	CToolTipTextProviderList * m_list;
public:
	CToolTipTextProviderItem( CToolTipTextProviderList * list );
	virtual ~CToolTipTextProviderItem();

	// separate methods for cases when the list can't be provided in constructor
	void Subscribe( CToolTipTextProviderList * list );
	void Unsubscribe();

	virtual BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );
};

class CToolTipTextProviderList: public CList < CToolTipTextProviderItem *, CToolTipTextProviderItem * >
{
public:
	CToolTipTextProviderList();
	virtual ~CToolTipTextProviderList();

	void AddProvider( CToolTipTextProviderItem * item );
	void RemoveProvider( CToolTipTextProviderItem * item );

	BOOL OnGetToolTipText( CMFCToolBarButton * pButton, CString & strTTText );
};
