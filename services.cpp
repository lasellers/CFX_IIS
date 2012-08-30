/*#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif*/

#include "stdafx.h"		// Standard MFC libraries
#include "activeds.h"
#include "afxpriv.h"	// MFC Unicode conversion macros

#include "cfx.h" //..\..\classes\cfx.h"		// CFX Custom Tag API

#include "basetsd.h"

#include "lmaccess.h" // UNLEN LM20_PWLEN

#include "wchar.h"

// bstr
#include <objidl.h>
#include <comdef.h>

// adsi
#include <iads.h>
#include <adshlp.h>

//
#include <stdlib.h>

#include <adssts.h>

// errs
#include "winerror.h"
#include "lmerr.h"

//safearray
#include "oleauto.h"

//
#include "common.h" //..\..\classes\common.h"
// about
//#include "about.h" //..\..\classes\about.h"
// ihtkpassword
//#include "ihtkpassword.h" //a..\..\classes\ihtkpassword.h"
//general error reporting
#include "err.h" //..\..\classes\err.h"
//com error reporting
#include "com.h" //..\..\classes\com.h"
#include "utility.h"
#include "support.h"


// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////

#include "services.h"

// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////

// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////


//
// required: server
// optional: computer directorytype
void StopStartPauseContinueNT( CCFXRequest* pRequest, CString op )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: StopStartPauseContinueNT<br>");
#endif

	// required
	CString strinServer = pRequest->GetAttribute( "Server" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	CString strinServiceType = pRequest->GetAttribute( "ServiceType" );
	pRequest->SetVariable( "IISServiceType", strinServiceType );
	
	CString svc;
	CString svckeytype;
	if(strinServiceType.CompareNoCase("FTP")==0)
	{
		svc="MSFTPSVC";
		svckeytype="IisFtpServer";
	}
	else if (strinServiceType.CompareNoCase("Web")==0)
	{
		svc="W3SVC";
		svckeytype="IisWebServer";
	}
	else if (strinServiceType.CompareNoCase("SMTP")==0)
	{
		svc="SMTPSVC";
		svckeytype="IIsSmtpServer";
	}
	
	// min parms
	if(	strinServer.IsEmpty() || strinServiceType.IsEmpty() )
	{
		log_ARGS(op,"SERVER and SERVICETYPE");
		return;
	}
	
	// optional
	CString strinComputer = pRequest->GetAttribute( "Computer" );
	pRequest->SetVariable( "IISComputer", strinComputer );
	CString computer;
	if(	strinComputer.IsEmpty() )
		computer="LocalHost";
	else
		computer=strinComputer;
	
	//
	HRESULT hr,hr2;
	CString pathe;
	
	//
	pathe.Format("IIS://%s/%s",computer,svc);
	
	IADsContainer *pContainer=NULL;
	hr = ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		IDispatch *pDisp=NULL;
		hr = pContainer->GetObject(A2W(svckeytype),A2W(strinServer),&pDisp);
		log_COMError(__LINE__,hr);
		
		//
		hr2 = pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
			//
			IADsServiceOperations *pSrvOp=NULL;
			hr = pDisp->QueryInterface(IID_IADsServiceOperations,(void**)&pSrvOp);
			log_COMError(__LINE__,hr);
			
			//
			hr2 = pDisp->Release();
			log_COMError(__LINE__,hr2);
			
			//
			if(SUCCEEDED(hr))
			{
				//
				long status=0;
				hr = pSrvOp->get_Status(&status);
				log_COMError(__LINE__,hr);
				
				//
				if(op.CompareNoCase("Stop")==0)
				{
					if(status == MD_SERVER_STATE_STARTED)
						pSrvOp->Stop();
				}
				else if(op.CompareNoCase("Start")==0)
				{
					if(status != MD_SERVER_STATE_STARTED)
						pSrvOp->Start();
				}
				else if(op.CompareNoCase("Pause")==0)
				{
					if(status != MD_SERVER_STATE_PAUSED)
						pSrvOp->Pause();
				}
				else if(op.CompareNoCase("Continue")==0)
				{
					if(status == MD_SERVER_STATE_PAUSED)
						pSrvOp->Continue();
				}
				
				//
				hr2 = pSrvOp->Release();
				log_COMError(__LINE__,hr2);
			}
		}
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: StopStartPauseContinueNT<br>");
#endif
}

// old trick to get the total number of lines of code
int getlines_services() { return __LINE__+1; }
