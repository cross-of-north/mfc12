// PropPage1.cpp : implementation file
//

#include "pch.h"
#include "afxdialogex.h"
#include "MyListCtrl.h"
#include "MFC12.h"
#include "PropPage1.h"


// PropPage1 dialog

IMPLEMENT_DYNAMIC(PropPage1, CMFCPropertyPage)

PropPage1::PropPage1( const int iResourceID, const CString stringID, CWnd* pParent /*=nullptr*/)
	: CMFCPropertyPage( iResourceID )
	, m_stringID( stringID )
{

}

PropPage1::~PropPage1()
{
}

void PropPage1::DoDataExchange(CDataExchange* pDX)
{
	CMFCPropertyPage::DoDataExchange(pDX);
	DDX_Control( pDX, IDC_LIST1, myListCtrl );
}


BEGIN_MESSAGE_MAP(PropPage1, CMFCPropertyPage)
END_MESSAGE_MAP()


// PropPage1 message handlers


BOOL PropPage1::OnInitDialog()
{
	CMFCPropertyPage::OnInitDialog();
	BOOL oKay { TRUE };

		// Load the large icons
	CImage gear, blocks, ball;

	HRESULT hr {};
	hr = gear.Load( L"res\\gear 32x32 8bpp.bmp" );
	
	hr = blocks.Load( L"res\\blocks 32x32 8bpp.bmp" );
	//hr = blocks.Load( L"res\\blocks 16x16 8bpp.bmp" ); // fails when the image is a different size
	//hr = blocks.Load( L"res\\Jasper.jpg" ); // fails when the image is a different image type
	
	hr = ball.Load( L"res\\ball 32x32 8bpp.bmp" );

	// Create the image list
	ctrllistItems.Create( 32, 32, ILC_COLOR32, 8, 1 );

	int imgNdx { -1 };
	imgNdx = ctrllistItems.Add( reinterpret_cast< CBitmap* >( &gear ), TRANSPARENT );
	int cnt1 = ctrllistItems.GetImageCount();
	IMAGEINFO imgInfo1;
	oKay = ctrllistItems.GetImageInfo( 0, &imgInfo1 );

	imgNdx = ctrllistItems.Add( reinterpret_cast< CBitmap* >( &blocks ), TRANSPARENT );
	int cnt2 = ctrllistItems.GetImageCount();
	IMAGEINFO imgInfo2;
	oKay = ctrllistItems.GetImageInfo( 0, &imgInfo2 );

	imgNdx = ctrllistItems.Add( reinterpret_cast< CBitmap* >( &ball ), TRANSPARENT );
	int cnt3 = ctrllistItems.GetImageCount();
	IMAGEINFO imgInfo3;
	oKay = ctrllistItems.GetImageInfo( 0, &imgInfo3 );

	// Count the number of images in the image list
	int cnt = ctrllistItems.GetImageCount();
	ASSERT( cnt == 3 ); // fail1 when loading 'blocks 16x16 8bpp.bmp: cnt ==2 and should be 3; fail2 when loading 'Jasper.jpg': cnt == 22 and should be 3

	// Set the image list
	auto myImgList = myListCtrl.SetImageList( &ctrllistItems, LVSIL_NORMAL );

	// Insert the ctrl list items
	LVITEM lvi;
	lvi.mask = LVIF_TEXT | LVIF_IMAGE;
	
	// item 1
	lvi.iItem = 0; // zero based
	lvi.iSubItem = 0; // zero to insert item
	lvi.pszText = const_cast< LPWSTR >( L"Gear item" );
	lvi.iImage = 0; // gear image
	int itemNdx1 = myListCtrl.InsertItem( &lvi );


	// item 2
	lvi.iItem = 1;
	lvi.iSubItem = 0; // zero to insert item
	lvi.pszText = const_cast< LPWSTR >( L"Blocks item" );
	lvi.iImage = 1; // blocks image
	int itemNdx2 = myListCtrl.InsertItem( &lvi );

	// item 3
	lvi.iItem = 2; 
	lvi.iSubItem = 0; // zero to insert item
	lvi.pszText = const_cast< LPWSTR >( L"Ball item" );
	lvi.iImage = 2; // ball image
	int itemNdx3 = myListCtrl.InsertItem( &lvi );

	
	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
