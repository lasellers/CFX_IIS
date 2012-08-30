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

#include "nntpserver.h"

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
// nntp server

//
//
void NNTPServersNT( CCFXRequest* pRequest )
{
	/*
	USES_CONVERSION;

#ifdef _DEBUG
	CString tmp;
	pRequest->Write("DEBUGGING: ENTRY: NNTPServersNT<br>");
#endif
	
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
	int iState = pColumns->AddString( "State" );
	int iDescription = pColumns->AddString( "Description" );
	
	int iBindings = pColumns->AddString( "Bindings" );
	
	int iVersion = pColumns->AddString( "Version" );
	
	int iMessageSizeLimit = pColumns->AddString( "MessageSizeLimit" );
	int iSessionSizeLimit = pColumns->AddString( "SessionSizeLimit" );
	int iMessagesPerConnectionLimit = pColumns->AddString( "MessagesPerConnectionLimit" );
	int iRecipientsPerMessageLimit = pColumns->AddString( "RecipientsPerMessageLimit" );
	
	int iBadMailDirectory = pColumns->AddString( "BadMailDirectory" );
	
	int iNonDeliveryMailTo = pColumns->AddString( "NonDeliveryMailTo" );
	int iBadMailTo = pColumns->AddString( "BadMailTo" );
	
	int iRetryInterval = pColumns->AddString( "RetryInterval" );
	
	int iDelayNotification = pColumns->AddString( "DelayNotification" );
	int iExpirationTimeout = pColumns->AddString( "ExpirationTimeout" );
	
	int iHopCount = pColumns->AddString( "HopCount" );
	int iMasqueradeDomain = pColumns->AddString( "MasqueradeDomain" );
	int iFullyQualifiedDomainName = pColumns->AddString( "FullyQualifiedDomainName" );
	int iSmartHost = pColumns->AddString( "SmartHost" );
	int iEnableReverseDnsLookup = pColumns->AddString( "EnableReverseDnsLookup" );
	CCFXQuery* pQuery = pRequest->AddQuery( get_query_variable(), pColumns );
	
	//
    HRESULT hr,hr2,hrw;
	CString pathe;
	
	//
	pathe.Format("IIS://%s/SmtpSVC",computer);

#ifdef _DEBUG
	tmp.Format("DEBUGGING: Trying container %s...<br>",pathe);
	pRequest->Write(tmp);
#endif
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: success<br>");
#endif

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: enumerating container...<br>");
#endif

		//
		IEnumVARIANT *pEnum=NULL;
		hrw = ADsBuildEnumerator(pContainer, &pEnum); 
		log_COMError(__LINE__,hrw);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hrw))
		{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: success<br>");
#endif
			//
			VARIANT var;
			ULONG ulElements=1;
			
			//
			while ( (SUCCEEDED(hrw)) && (ulElements==1) )
			{
#ifdef _DEBUG
	tmp.Format("DEBUGGING: enumerating %d elements...<br>",ulElements);
	pRequest->Write(tmp);
#endif
				//
				hrw = ADsEnumerateNext(pEnum, 1, &var, &ulElements);
				log_COMError(__LINE__,hrw);
				
				//
				if(SUCCEEDED(hrw) && (ulElements==1) )
				{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: success<br>");
#endif
					//
					CString strServer;
					CString strState;
					CString strDescription;
					
					CString strBindings;
					
					CString strVersion;
					
					CString strMessageSizeLimit;
					CString strSessionSizeLimit;
					CString strMessagesPerConnectionLimit;
					CString strRecipientsPerMessageLimit;
					
					CString strBadMailDirectory;
					
					CString strNonDeliveryMailTo;
					CString strBadMailTo;
					
					CString strRetryInterval;
					
					CString strDelayNotification;
					CString strExpirationTimeout;
					
					CString strHopCount;
					CString strMasqueradeDomain;
					CString strFullyQualifiedDomainName;
					CString strSmartHost;
					CString strEnableReverseDnsLookup;
					
					//
					IADs *pADs=(IADs*)var.pdispVal;
					
					//
					BSTR bstrName;
					hr2 = pADs->get_Name(&bstrName);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strServer=bstrName;
					
					CString strClass;
					BSTR bstrClass;
					hr2 = pADs->get_Class(&bstrClass);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strClass=bstrClass;
					
#ifdef _DEBUG
	tmp.Format("DEBUGGING: server %s class %s...<br>",strServer,strClass);
	pRequest->Write(tmp);
#endif
					//
					if(strClass.Compare("IIsSmtpServer")==0)
					{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: success<br>");
#endif
						//
						VARIANT var;
						VariantInit(&var);
						
						//
						hr2 = pADs->Get(L"ServerState",&var);
						if(SUCCEEDED(hr2))
						{
							switch(V_INT(&var))
							{
							case MD_SERVER_STATE_STARTING: strState="Starting"; break;
							case MD_SERVER_STATE_STARTED: strState="Started"; break;
							case MD_SERVER_STATE_STOPPING: strState="Stopping"; break;
							case MD_SERVER_STATE_STOPPED: strState="Stopped"; break;
							case MD_SERVER_STATE_PAUSING: strState="Pausing"; break;
							case MD_SERVER_STATE_PAUSED: strState="Paused"; break;
							case MD_SERVER_STATE_CONTINUING: strState="Continuing"; break;
							default: strState="Unknown";
							}
						}
						VariantClear(&var);
						
						hr2 = pADs->Get(L"ServerComment",&var);
						if(SUCCEEDED(hr2)) strDescription = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"ServerBindings",&var);
						if(SUCCEEDED(hr2))
						{
							VARIANTARRAYtoCString(var,strBindings,pRequest,';');
						}
						VariantClear(&var);
						
						hr2 = pADs->Get(L"SmtpServiceVersion",&var);
						if(SUCCEEDED(hr2)) strVersion.Format("%d",V_INT(&var));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"MaxMessageSize",&var);
						if(SUCCEEDED(hr2)) strMessageSizeLimit.Format("%d",V_INT(&var));
						VariantClear(&var);
						hr2 = pADs->Get(L"MaxSessionSize",&var);
						if(SUCCEEDED(hr2)) strSessionSizeLimit.Format("%d",V_INT(&var));
						VariantClear(&var);
						hr2 = pADs->Get(L"MaxBatchedMessages",&var);
						if(SUCCEEDED(hr2)) strMessagesPerConnectionLimit.Format("%d",V_INT(&var));
						VariantClear(&var);
						hr2 = pADs->Get(L"MaxRecipients",&var);
						if(SUCCEEDED(hr2)) strRecipientsPerMessageLimit.Format("%d",V_INT(&var));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"BadMailDirectory",&var);
						if(SUCCEEDED(hr2)) strBadMailDirectory = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"SendNdrTo",&var);
						if(SUCCEEDED(hr2)) strNonDeliveryMailTo = V_BSTR(&var) ;
						VariantClear(&var);
						hr2 = pADs->Get(L"SendBadTo",&var);
						if(SUCCEEDED(hr2)) strBadMailTo = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"SmtpRemoteProgressiveRetry",&var);
						if(SUCCEEDED(hr2)) strRetryInterval = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"SmtpLocalDelayExpireMinutes",&var);
						if(SUCCEEDED(hr2)) strDelayNotification.Format("%d",V_INT(&var));
						VariantClear(&var);
						hr2 = pADs->Get(L"SmtpLocalNDRExpireMinutes",&var);
						if(SUCCEEDED(hr2)) strExpirationTimeout.Format("%d",V_INT(&var));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"HopCount",&var);
						if(SUCCEEDED(hr2)) strHopCount.Format("%d",V_INT(&var));
						VariantClear(&var);
						hr2 = pADs->Get(L"MasqueradeDomain",&var);
						if(SUCCEEDED(hr2)) strMasqueradeDomain = V_BSTR(&var) ;
						VariantClear(&var);
						hr2 = pADs->Get(L"FullyQualifiedDomainName",&var);
						if(SUCCEEDED(hr2)) strFullyQualifiedDomainName = V_BSTR(&var) ;
						VariantClear(&var);
						hr2 = pADs->Get(L"SmartHost",&var);
						if(SUCCEEDED(hr2)) strSmartHost = V_BSTR(&var) ;
						VariantClear(&var);
						hr2 = pADs->Get(L"EnableReverseDnsLookup",&var);
						if(SUCCEEDED(hr2)) strEnableReverseDnsLookup.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						//
						int iRow = pQuery->AddRow();
						pQuery->SetData( iRow, iServer, strServer );
						pQuery->SetData( iRow, iState, strState );
						pQuery->SetData( iRow, iDescription, strDescription );
						
						pQuery->SetData( iRow, iBindings, strBindings );
						
						pQuery->SetData( iRow, iVersion, strVersion );
						
						pQuery->SetData( iRow, iMessageSizeLimit, strMessageSizeLimit );
						pQuery->SetData( iRow, iSessionSizeLimit, strSessionSizeLimit );
						pQuery->SetData( iRow, iMessagesPerConnectionLimit, strMessagesPerConnectionLimit );
						pQuery->SetData( iRow, iRecipientsPerMessageLimit, strRecipientsPerMessageLimit );
						
						pQuery->SetData( iRow, iBadMailDirectory, strBadMailDirectory );
						
						pQuery->SetData( iRow, iNonDeliveryMailTo, strNonDeliveryMailTo );
						pQuery->SetData( iRow, iBadMailTo, strBadMailTo );
						
						pQuery->SetData( iRow, iRetryInterval, strRetryInterval );
						
						pQuery->SetData( iRow, iDelayNotification, strDelayNotification );
						pQuery->SetData( iRow, iExpirationTimeout, strExpirationTimeout );
						
						pQuery->SetData( iRow, iHopCount, strHopCount );
						pQuery->SetData( iRow, iMasqueradeDomain, strMasqueradeDomain );
						pQuery->SetData( iRow, iFullyQualifiedDomainName, strFullyQualifiedDomainName );
						pQuery->SetData( iRow, iSmartHost, strSmartHost );
						pQuery->SetData( iRow, iEnableReverseDnsLookup, strEnableReverseDnsLookup );
					}
					
					//
					hr2=pADs->Release();
					log_COMError(__LINE__,hr2);
					
					//
					SysFreeString(bstrClass);
					SysFreeString(bstrName);
				}

			}
		}
		//
		hr2 = ADsFreeEnumerator(pEnum);
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: SMTPServersNT<br>");
#endif
	*/
}

// old trick to get the total number of lines of code
int getlines_nntpserver() { return __LINE__+1; }
