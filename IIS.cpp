// IIS.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include "IIS.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CIISApp

BEGIN_MESSAGE_MAP(CIISApp, CWinApp)
	//{{AFX_MSG_MAP(CIISApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CIISApp construction

CIISApp::CIISApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CIISApp object

CIISApp theApp;
