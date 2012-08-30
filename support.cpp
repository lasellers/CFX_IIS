/* #ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif */

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

CString umQuery="";
void set_query_variable(CString str)
{
	umQuery=str;
}
CString get_query_variable()
{
	return umQuery;
}

// computer: LocalHost, etc
// service: MSFTPSVC W3SVC
// class: IIsFtpServer IIsWebServer
int NextServer( CCFXRequest* pRequest, CString strinComputer, CString strinService)
{
	USES_CONVERSION;

#ifdef _DEBUG
	CString tmp;
	pRequest->Write("<ul>DEBUGGING: ENTRY: NextServer...<br>");
#endif
	
	//
	CString computer;
	if(	strinComputer.IsEmpty() )
		computer="LocalHost";
	else
		computer=strinComputer;
	
	//
	int serverno=0;
	
	//
	HRESULT hr,hr2;
	CString pathe;
	
	//
	pathe.Format("IIS://%s/%s", computer,strinService);

#ifdef _DEBUG
	tmp.Format("DEBUGGING: Trying to attach to container %s...<br>",pathe);
	pRequest->Write(tmp);
#endif
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr) && pContainer)
	{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Container attached...<br>");
#endif
		//
		HRESULT hrw;
		IEnumVARIANT *pEnum=NULL;
		hrw = ADsBuildEnumerator(pContainer, &pEnum); 
		log_COMError(__LINE__,hrw);

#ifdef _DEBUG
	tmp.Format("DEBUGGING: Trying to enumerate...<br>");
	pRequest->Write(tmp);
#endif
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hrw))
		{
#ifdef _DEBUG
	tmp.Format("DEBUGGING: Attached...<br>",tmp);
	pRequest->Write(tmp);
#endif
			//
			VARIANT var;
			ULONG ulElements=1;
			
			//
			while ( (SUCCEEDED(hrw)) && (ulElements==1) )
			{
#ifdef _DEBUG
	tmp.Format("<ul>DEBUGGING: Enumerating %d elements...<br>",ulElements);
	pRequest->Write(tmp);
#endif
				//
				hrw = ADsEnumerateNext(pEnum, 1, &var, &ulElements);
				log_COMError(__LINE__,hrw);
				
				//
				if( (SUCCEEDED(hrw)) && (ulElements==1) )
				{
#ifdef _DEBUG
	tmp.Format("DEBUGGING: success...<br>");
	pRequest->Write(tmp);
#endif
					//
					IADs *pADs=(IADs*)var.pdispVal;
					
					//
					BSTR bstrName;
					hr = pADs->get_Name(&bstrName);
#ifdef _DEBUG
	CString strServer=bstrName;
	tmp.Format("DEBUGGING: server name %s...<br>",strServer);
	pRequest->Write(tmp);
#endif
					log_COMError(__LINE__,hr);
					int sno=0;
					if(SUCCEEDED(hr))
					{
						CString strNo=bstrName;
						sno=atoi(strNo);
						if(sno>serverno)
						{
							serverno=sno;
#ifdef _DEBUG
	tmp.Format("DEBUGGING: +new highest server no %d...<br>",serverno);
	pRequest->Write(tmp);
#endif
						}
					}
					
					// ?
					hr2=pADs->Release();
					log_COMError(__LINE__,hr2);
					
					//
					SysFreeString(bstrName);
				}
#ifdef _DEBUG
	pRequest->Write("</ul><br>");
#endif
			}
		}
		
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Releasing enumerator...<br>");
#endif
		//		
		hr2 = ADsFreeEnumerator(pEnum);
		log_COMError(__LINE__,hr2);
	}
	
#ifdef _DEBUG
	tmp.Format("DEBUGGING: results: next free server is %d...<br>",(serverno+1));
	pRequest->Write(tmp);
#endif

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: NextServer...<br></ul>");
#endif

	return serverno+1;
}
