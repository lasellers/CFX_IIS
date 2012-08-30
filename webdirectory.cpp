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

#include "webdirectory.h"

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

// web directory

//
//
// Required: Server
// optional: computer
void WebDirectoriesNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: WebDirectoriesNT<br>");
#endif

	// required
	CString strinServer = pRequest->GetAttribute( "Server" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	// min parms
	if(	strinServer.IsEmpty() )
	{
		log_ARGS("WebDirectories","SERVER");
		return;
	}
	
	// optional
	CString strinComputer = pRequest->GetAttribute( "COMPUTER" );
	pRequest->SetVariable( "IISComputer", strinComputer );
	CString computer;
	if(	strinComputer.IsEmpty() )
		computer="LocalHost";
	else
		computer=strinComputer;
	
	// fields
	CCFXStringSet* pColumns = pRequest->CreateStringSet();
	int iServer = pColumns->AddString( "Server" );
	int iDirectoryType = pColumns->AddString( "DirectoryType" );
	int iDescription = pColumns->AddString( "Description" );
	
	int iAnonymousUsername = pColumns->AddString( "AnonymousUsername" );
	int iAnonymousPassword = pColumns->AddString( "AnonymousPassword" );
	
	int iAccessRead = pColumns->AddString( "AccessRead" );
	int iAccessWrite = pColumns->AddString( "AccessWrite" );
	int iAccessExecute = pColumns->AddString( "AccessExecute" );
	int iAccessScript = pColumns->AddString( "AccessScript" );
	
	int iAuthBasic = pColumns->AddString( "AuthBasic" );
	int iAuthAnonymous = pColumns->AddString( "AuthAnonymous" );
	int iAuthNTLM = pColumns->AddString( "AuthNTLM" );
	
	int iDirectoryBrowsingAllowed = pColumns->AddString( "DirectoryBrowsingAllowed" );
	int iContentIndexed = pColumns->AddString( "ContentIndexed" );
	
	int iEnableDefaultDoc = pColumns->AddString( "EnableDefaultDoc" );
	int iDefaultDoc = pColumns->AddString( "DefaultDoc" );
	
	int iEnableDocFooter = pColumns->AddString( "EnableDocFooter" );
	int iDocFooter = pColumns->AddString( "DocFooter" );
	
	/*	dontlog */
	int iPath = pColumns->AddString( "Path" );
	int iDefaultLogonDomain = pColumns->AddString( "DefaultLogonDomain" );
	int iUNCUsername = pColumns->AddString( "UNCUsername" );
	int iUNCPassword = pColumns->AddString( "UNCPassword" );
	
	int iCustomErrors = pColumns->AddString( "CustomErrors" );
	int iHTTPHeaders = pColumns->AddString( "HTTPHeaders" );
	int iHTTPRedirect = pColumns->AddString( "HTTPRedirect" );

	CCFXQuery* pQuery = pRequest->AddQuery( get_query_variable(), pColumns );
	
	//
    HRESULT hr,hr2;
	CString pathe;
	
	//
	pathe.Format("IIS://%s/W3SVC/%s/Root",computer,strinServer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		IEnumVARIANT *pEnum=NULL;
		hr = ADsBuildEnumerator(pContainer, &pEnum); 
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
			//
			ULONG ulElements=1;
			VARIANT var;
			
			//
			while ( (SUCCEEDED(hr)) && (ulElements==1) )
			{
				//
				hr = ADsEnumerateNext(pEnum, 1, &var, &ulElements);
				log_COMError(__LINE__,hr);
				
				//
				if(SUCCEEDED(hr) && (ulElements==1) )
				{
					//
					IADs *pADs=(IADs*)var.pdispVal;
					
					//
					BSTR bstrClass;
					hr2 = pADs->get_Class(&bstrClass);
					log_COMError(__LINE__,hr2);
					CString strClass(bstrClass);
					
					//
					if(strClass.Compare("IIsWebVirtualDir")==0)
					{
						//
						CString strServer;
						
						CString strDescription;
						CString strDirectoryType;
						
						CString strAnonymousUsername;
						CString strAnonymousPassword;
						
						CString strAccessRead;
						CString strAccessWrite; 
						CString strAccessExecute; 
						CString strAccessScript; 
						
						CString strAuthBasic; 
						CString strAuthAnonymous; 
						CString strAuthNTLM; 
						
						CString strDirectoryBrowsingAllowed;
						CString strContentIndexed;
						
						CString strEnableDefaultDoc;
						CString strDefaultDoc;
						
						CString strEnableDocFooter;
						CString strDocFooter;
						
						CString strPath;
						CString strDefaultLogonDomain;
						CString strUNCUsername;
						CString strUNCPassword;
						
						CString strCustomErrors;
						CString strHTTPHeaders;
						CString strHTTPRedirect;
						
						//
						BSTR bstrServer;
						hr2 = pADs->get_Name(&bstrServer);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strServer=bstrServer;
						
						BSTR bstrDescription;
						hr2 = pADs->get_Name(&bstrDescription);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDescription=bstrDescription;
						
						//
						VARIANT var;
						VariantInit(&var);
						
						//
						hr2 = pADs->Get(L"AnonymousUserName",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAnonymousUsername = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AnonymousUserPass",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAnonymousPassword = V_BSTR(&var) ;
						VariantClear(&var);
						
						if(strClass.Compare("IIsWebVirtualDir")==0)
							strDirectoryType="Virtual";
						else if (strClass.Compare("IIsWebDirectory")==0)
							strDirectoryType="Subdirectory";
						
						hr2 = pADs->Get(L"AccessRead",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessRead.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AccessWrite",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessWrite.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AccessExecute",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessExecute.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AccessScript",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessScript.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AuthBasic",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAuthBasic.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AuthAnonymous",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAuthAnonymous.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AuthNTLM",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAuthNTLM.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"EnableDirBrowsing",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDirectoryBrowsingAllowed.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"ContentIndexed",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strContentIndexed.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"EnableDefaultDoc",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strEnableDefaultDoc.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"DefaultDoc",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDefaultDoc = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"EnableDocFooter",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strEnableDocFooter.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"DefaultDocFooter",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDocFooter = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"Path",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strPath = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"DefaultLogonDomain",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDefaultLogonDomain = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"UNCUserName",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strUNCUsername = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"UNCPassword",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strUNCPassword = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"HTTPErrors",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strCustomErrors,pRequest,';');
						VariantClear(&var);
						
						hr2 = pADs->Get(L"HTTPCustomHeaders",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strHTTPHeaders,pRequest,',');
						VariantClear(&var);
						
						hr2 = pADs->Get(L"HttpRedirect",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strHTTPRedirect = V_BSTR(&var) ;
						VariantClear(&var);
						
						//
						int iRow = pQuery->AddRow();
						pQuery->SetData( iRow, iServer, strServer );
						pQuery->SetData( iRow, iDirectoryType, strDirectoryType );
						pQuery->SetData( iRow, iDescription, strDescription );
						
						pQuery->SetData( iRow, iAnonymousUsername, strAnonymousUsername );
						pQuery->SetData( iRow, iAnonymousPassword, strAnonymousPassword );
						
						pQuery->SetData( iRow, iAccessRead, strAccessRead );
						pQuery->SetData( iRow, iAccessWrite, strAccessWrite );
						pQuery->SetData( iRow, iAccessExecute, strAccessExecute );
						pQuery->SetData( iRow, iAccessScript, strAccessScript );
						
						pQuery->SetData( iRow, iAuthBasic, strAuthBasic );
						pQuery->SetData( iRow, iAuthAnonymous, strAuthAnonymous );
						pQuery->SetData( iRow, iAuthNTLM, strAuthNTLM );
						
						pQuery->SetData( iRow, iDirectoryBrowsingAllowed, strDirectoryBrowsingAllowed );
						
						pQuery->SetData( iRow, iEnableDefaultDoc, strEnableDefaultDoc );
						pQuery->SetData( iRow, iDefaultDoc, strDefaultDoc );
						
						pQuery->SetData( iRow, iEnableDocFooter, strEnableDocFooter );
						pQuery->SetData( iRow, iDocFooter, strDocFooter );
						
						pQuery->SetData( iRow, iContentIndexed, strContentIndexed );
						pQuery->SetData( iRow, iPath, strPath );
						pQuery->SetData( iRow, iDefaultLogonDomain, strDefaultLogonDomain );
						pQuery->SetData( iRow, iUNCUsername, strUNCUsername );
						pQuery->SetData( iRow, iUNCPassword, strUNCPassword );
						
						pQuery->SetData( iRow, iCustomErrors, strCustomErrors );
						pQuery->SetData( iRow, iHTTPHeaders, strHTTPHeaders );
						pQuery->SetData( iRow, iHTTPRedirect, strHTTPRedirect );
						
						//
						SysFreeString(bstrDescription);
						SysFreeString(bstrServer);
					}
					
					//
					SysFreeString(bstrClass);
					
					//
					hr2=pADs->Release();
					log_COMError(__LINE__,hr2);
				}
			}
		}
		
		//		
		hr2 = ADsFreeEnumerator(pEnum);
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: WebDirectoriesNT<br>");
#endif
}




//
// required: server directory path
// optional:computer uncusername uncpassword
void AddWebDirectoryNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: AddWebDirectoryNT<br>");
#endif

	// required
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	CString strinDirectory = pRequest->GetAttribute( "DIRECTORY" );
	pRequest->SetVariable( "IISDirectory", strinDirectory );
	
	CString strinPath = pRequest->GetAttribute( "PATH" );
	pRequest->SetVariable( "IISPath", strinPath );
	
	// min parms
	if(	strinServer.IsEmpty() || strinDirectory.IsEmpty() || strinPath.IsEmpty() )
	{
		log_ARGS("AddWebDirectory","SERVER, DIRECTORY and PATH");
		return;
	}
	
	// optional
	CString strinComputer = pRequest->GetAttribute( "COMPUTER" );
	pRequest->SetVariable( "IISComputer", strinComputer );
	CString computer;
	if(	strinComputer.IsEmpty() )
		computer="LocalHost";
	else
		computer=strinComputer;
	
	CString strinAnonymousUsername = pRequest->GetAttribute( "AnonymousUsername" );
	CString strinAnonymousPassword = pRequest->GetAttribute( "AnonymousPassword" );
	
	CString strinAuthBasic = pRequest->GetAttribute( "AuthBasic" );
	CString strinAuthAnonymous = pRequest->GetAttribute( "AuthAnonymous" );
	CString strinAuthNTLM = pRequest->GetAttribute( "AuthNTLM" );
	
	CString strinAccessRead = pRequest->GetAttribute( "AccessRead" );
	CString strinAccessWrite = pRequest->GetAttribute( "AccessWrite" );
	CString strinAccessExecute = pRequest->GetAttribute( "AccessExecute" );
	CString strinAccessScript = pRequest->GetAttribute( "AccessScript" );
	
	CString strinDirectoryBrowsingAllowed = pRequest->GetAttribute( "DirectoryBrowsingAllowed" );
	
	CString strinEnableDefaultDoc = pRequest->GetAttribute( "EnableDefaultDoc" );
	CString strinDefaultDoc = pRequest->GetAttribute( "DefaultDoc" );
	
	CString strinEnableDocFooter = pRequest->GetAttribute( "EnableDocFooter" );
	CString strinDocFooter = pRequest->GetAttribute( "DocFooter" );
	
	CString strinContentIndexed = pRequest->GetAttribute( "ContentIndexed" );
	CString strinDefaultLogonDomain = pRequest->GetAttribute( "DefaultLogonDomain" );
	CString strinUNCUsername = pRequest->GetAttribute( "UNCUsername" );
	CString strinUNCPassword = pRequest->GetAttribute( "UNCPassword" );
	CString strinCustomErrors = pRequest->GetAttribute( "CustomErrors" );
	CString strinHTTPHeaders = pRequest->GetAttribute( "HTTPHeaders" );
	CString strinHTTPRedirect = pRequest->GetAttribute( "HTTPRedirect" );
	
	//
    HRESULT hr,hr2;
	CString pathe;
	
	//
	pathe.Format("IIS://%s/W3SVC/%s/Root",computer,strinServer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		IDispatch *pDisp=NULL;
		hr = pContainer->Create(L"IIsWebVirtualDir", A2W(strinDirectory), &pDisp );
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if( SUCCEEDED(hr) )
		{
			//
			IADs *pADs=NULL;
			hr = pDisp->QueryInterface( IID_IADs, (void**)&pADs );
			log_COMError(__LINE__,hr);
			
			//
			hr2=pADs->SetInfo();
			log_COMError(__LINE__,hr2);
			
			//
			hr2=pDisp->Release();
			log_COMError(__LINE__,hr2);
			
			//
			if ( SUCCEEDED(hr) )
			{
				VARIANT var;
				VariantInit(&var);
				
				//set rquired properties
				V_BSTR(&var) = SysAllocString(L"IIsWebVirtualDir");
				V_VT(&var) = VT_BSTR;
				hr2=pADs->Put(L"KeyType",var);
				log_COMError(__LINE__,hr2);
				VariantClear(&var);
				
				V_BSTR(&var) = SysAllocString(A2W(strinPath));
				V_VT(&var) = VT_BSTR;
				hr2=pADs->Put(L"Path",var);
				log_COMError(__LINE__,hr2);
				VariantClear(&var);
				
				// optionals
				if(!strinAnonymousUsername.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinAnonymousUsername));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"AnonymousUserName",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAnonymousPassword.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinAnonymousPassword));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"AnonymousUserPass",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAuthBasic.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAuthBasic);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AuthBasic",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAuthAnonymous.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAuthAnonymous);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AuthAnonymous",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAuthNTLM.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAuthNTLM);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AuthNTLM",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAccessRead.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAccessRead);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AccessRead",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAccessWrite.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAccessWrite);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AccessWrite",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAccessExecute.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAccessExecute);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AccessExecute",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAccessScript.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAccessScript);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AccessScript",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinDirectoryBrowsingAllowed.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinDirectoryBrowsingAllowed);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"EnableDirBrowsing",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinContentIndexed.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinContentIndexed);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"ContentIndexed",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinEnableDefaultDoc.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinEnableDefaultDoc);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"EnableDefaultDoc",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinDefaultDoc.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinDefaultDoc));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"DefaultDoc",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinEnableDocFooter.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinEnableDocFooter);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"EnableDocFooter",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinDocFooter.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinDocFooter));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"DefaultDocFooter",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinDefaultLogonDomain.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinDefaultLogonDomain));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"DefaultLogonDomain",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinUNCUsername.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinUNCUsername));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"UNCUserName",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinUNCPassword.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinUNCPassword));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"UNCPassword",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinCustomErrors.IsEmpty())
				{
					CStringtoVARIANTARRAY(strinCustomErrors,&var,pRequest,';');
					hr2=pADs->Put(L"HttpErrors",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinHTTPHeaders.IsEmpty())
				{
					CStringtoVARIANTARRAY(strinHTTPHeaders,&var,pRequest,',');
					hr2=pADs->Put(L"HTTPCustomHeaders",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinHTTPRedirect.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinHTTPRedirect));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"HTTPRedirect",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				// write it out
				hr2=pADs->SetInfo();
				log_COMError(__LINE__,hr2);
				
				//
				hr2=pADs->Release();
				log_COMError(__LINE__,hr2);
			}
		}
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: AddWebDirectoryNT<br>");
#endif
}


//
// required: server directory
// optional: computer
void DeleteWebDirectoryNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: DeleteWebDirectoryNT<br>");
#endif

	// required
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	CString strinDirectory = pRequest->GetAttribute( "DIRECTORY" );
	pRequest->SetVariable( "IISDirectory", strinDirectory );
	
	// min parms
	if(	strinServer.IsEmpty() || strinDirectory.IsEmpty() )
	{
		log_ARGS("DeleteWebDirectory","SERVER and DIRECTORY");
		return;
	}
	
	// optional
	CString strinComputer = pRequest->GetAttribute( "COMPUTER" );
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
	pathe.Format("IIS://%s/W3SVC/%s/Root",computer,strinServer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		hr=pContainer->Delete(L"IIsWebVirtualDir",A2W(strinDirectory));
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: DeleteWebDirectoryNT<br>");
#endif
}

// old trick to get the total number of lines of code
int getlines_webdirectory() { return __LINE__+1; }
