// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface 
// (the "Fluent UI") and is provided only as referential material to supplement the 
// Microsoft Foundation Classes Reference and related electronic documentation 
// included with the MFC C++ library software.  
// License terms to copy, use or distribute the Fluent UI are available separately.  
// To learn more about our Fluent UI licensing program, please visit 
// https://go.microsoft.com/fwlink/?LinkId=238214.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// ChildFrm.cpp : implementation of the ChildFrame class
//

#include "pch.h"
#include "framework.h"
#include "MFC12.h"

#include "ChildFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// ChildFrame

IMPLEMENT_DYNCREATE(ChildFrame, CMDIChildWndEx)

BEGIN_MESSAGE_MAP(ChildFrame, CMDIChildWndEx)
	ON_COMMAND(ID_FILE_PRINT, &ChildFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &ChildFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &ChildFrame::OnFilePrintPreview)
	ON_UPDATE_COMMAND_UI(ID_FILE_PRINT_PREVIEW, &ChildFrame::OnUpdateFilePrintPreview)
END_MESSAGE_MAP()

// ChildFrame construction/destruction

ChildFrame::ChildFrame() noexcept
{
	// TODO: add member initialization code here
}

ChildFrame::~ChildFrame()
{
}


BOOL ChildFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying the CREATESTRUCT cs
	if( !CMDIChildWndEx::PreCreateWindow(cs) )
		return FALSE;

	return TRUE;
}

// ChildFrame diagnostics

#ifdef _DEBUG
void ChildFrame::AssertValid() const
{
	CMDIChildWndEx::AssertValid();
}

void ChildFrame::Dump(CDumpContext& dc) const
{
	CMDIChildWndEx::Dump(dc);
}
#endif //_DEBUG

// ChildFrame message handlers

void ChildFrame::OnFilePrint()
{
	if (m_dockManager.IsPrintPreviewValid())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_PRINT);
	}
}

void ChildFrame::OnFilePrintPreview()
{
	if (m_dockManager.IsPrintPreviewValid())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_CLOSE);  // force Print Preview mode closed
	}
}

void ChildFrame::OnUpdateFilePrintPreview(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(m_dockManager.IsPrintPreviewValid());
}
