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

#include "webserver.h"

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

// web server

//
//
void WebServersNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	CString tmp;
	pRequest->Write("DEBUGGING: ENTRY: WebServerNT<br>");
#endif

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Establishing SQL fields...<br>");
#endif

	// fields
	CCFXStringSet* pColumns = pRequest->CreateStringSet();
	int iServer = pColumns->AddString( "Server" );
	int iState = pColumns->AddString( "State" );
	int iDescription = pColumns->AddString( "Description" );
	
	int iPath = pColumns->AddString( "Path" );
	
	int iAnonymousUsername = pColumns->AddString( "AnonymousUsername" );
	int iAnonymousPassword = pColumns->AddString( "AnonymousPassword" );
	
	int iBindings = pColumns->AddString( "Bindings" );
	int iSSLBindings = pColumns->AddString( "SSLBindings" );
	
	int iMaxConnections = pColumns->AddString( "MaxConnections" );
	int iConnectionTimeout = pColumns->AddString( "ConnectionTimeout" );
	int iDefaultLogonDomain = pColumns->AddString( "DefaultLogonDomain" );
	
	int iEnableDefaultDoc = pColumns->AddString( "EnableDefaultDoc" );
	int iDefaultDoc = pColumns->AddString( "DefaultDoc" );
	
	int iEnableDocFooter = pColumns->AddString( "EnableDocFooter" );
	int iDocFooter = pColumns->AddString( "DocFooter" );
	
	int iAccessRead = pColumns->AddString( "AccessRead" );
	int iAccessWrite = pColumns->AddString( "AccessWrite" );
	int iAccessExecute = pColumns->AddString( "AccessExecute" );
	int iAccessScript = pColumns->AddString( "AccessScript" );
	int iDirectoryBrowsingAllowed = pColumns->AddString( "DirectoryBrowsingAllowed" );
	
	int iCustomErrors = pColumns->AddString( "CustomErrors" );
	int iHTTPHeaders = pColumns->AddString( "HTTPHeaders" );
	int iHTTPRedirect = pColumns->AddString( "HTTPRedirect" );
	
	int iApplicationName = pColumns->AddString( "ApplicationName" );
	int iApplicationProtection = pColumns->AddString( "ApplicationProtection" );
	int iApplicationStartingPoint = pColumns->AddString( "ApplicationStartingPoint" );

	//
	int iEnableLogging = pColumns->AddString( "EnableLogging" );
	int iActiveLogFormat = pColumns->AddString( "ActiveLogFormat" );
	int iLogFileDirectory = pColumns->AddString( "LogFileDirectory" );
	int iNewLogTimePeriod = pColumns->AddString( "NewLogTimePeriod" );
	int iMB = pColumns->AddString( "MB" );
	int iUseLocalTime = pColumns->AddString( "UseLocalTime" );
	int iODBCDataSource = pColumns->AddString( "ODBCDataSource" );
	int iODBCTable = pColumns->AddString( "ODBCTable" );
	int iODBCUserName = pColumns->AddString( "ODBCUserName" );
	int iODBCPassword = pColumns->AddString( "ODBCPassword" );

	CCFXQuery* pQuery = pRequest->AddQuery( get_query_variable(), pColumns );
	
	// optional
	CString strinComputer = pRequest->GetAttribute( "COMPUTER" );
	pRequest->SetVariable( "IISComputer", strinComputer );
	CString computer;
	if(	strinComputer.IsEmpty() )
		computer="LocalHost";
	else
		computer=strinComputer;
	
	//
    HRESULT hr,hr2,hrr;
	CString pathe;

	
	// get logging module names
	CString logidIIS;
	CString logidW3C;
	CString logidNCSA;
	CString logidODBC;
    IADs *pADs=NULL;
	{
		VARIANT var;
		VariantInit(&var);

		// iis
		pADs=NULL;
		pathe.Format("IIS://%s/logging/Microsoft IIS Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidIIS = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// w3c
		pADs=NULL;
		pathe.Format("IIS://%s/logging/W3C Extended Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidW3C = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// ncsa
		pADs=NULL;
		pathe.Format("IIS://%s/logging/NCSA Common Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidNCSA = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// odbc logging
		pADs=NULL;
		pathe.Format("IIS://%s/logging/ODBC Logging",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidODBC = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
	}



	//
	pathe.Format("IIS://%s/W3SVC",computer);
#ifdef _DEBUG
	tmp.Format("DEBUGGING: container %s...<br>",pathe);
	pRequest->Write(tmp);
#endif
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: container attached...<br>");
#endif
		//
		IEnumVARIANT *pEnum;
		hr = ADsBuildEnumerator(pContainer, &pEnum); 
		log_COMError(__LINE__,hr);
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Preparing to enumerate container...<br>");
#endif		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Enumeration can start...<br>");
#endif
			//
			VARIANT var;
			ULONG ulElements=1;
			
			//
			while ( (SUCCEEDED(hr)) && (ulElements==1) )
			{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Enumerating...<br><ul>");
#endif
				//
				hr = ADsEnumerateNext(pEnum, 1, &var, &ulElements);
				log_COMError(__LINE__,hr);
#ifdef _DEBUG
	tmp.Format("DEBUGGING: enumerating %d elements...<br>",ulElements);
	pRequest->Write(tmp);
#endif
				
				//
				if(SUCCEEDED(hr) && (ulElements==1) )
				{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Enumeration succeeds...<br>");
#endif
					//
					IADs *pADs=(IADs*)var.pdispVal;
					
					//
					CString strServer;
					CString strState; //ServerState
					CString strDescription; //ServerComment;
					
					CString strPath;
					
					CString strAnonymousUsername;
					CString strAnonymousPassword;
					
					CString strBindings; //ServerBindings
					CString strSSLBindings; //SecureBindings
					
					CString strMaxConnections;
					CString strConnectionTimeout;
					CString strDefaultLogonDomain;
					
					CString strEnableDefaultDoc;
					CString strDefaultDoc;
					
					CString strEnableDocFooter;
					CString strDocFooter;
					
					CString strAccessRead;
					CString strAccessWrite; 
					CString strAccessExecute; 
					CString strAccessScript; 
					CString strDirectoryBrowsingAllowed; //EnableDirBrowsing
					
					CString strCustomErrors; //HTTPErrors
					CString strHTTPHeaders; //HTTPCustomHeaders
					CString strHTTPRedirect;
					
					CString strApplicationName;
					CString strApplicationProtection;
					CString strApplicationStartingPoint;

					//
					CString strEnableLogging;
					CString strActiveLogFormat;
					CString strLogFileDirectory;
					CString strNewLogTimePeriod;
					CString strMB;
					CString strUseLocalTime;
					CString strODBCDataSource;
					CString strODBCTable;
					CString strODBCUserName;
					CString strODBCPassword;
					
					//
					BSTR bstrName;
					hr2 = pADs->get_Name(&bstrName);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strServer=bstrName;
#ifdef _DEBUG
	tmp.Format("DEBUGGING: +server %s...<br>",strServer);
	pRequest->Write(tmp);
#endif
				
					//
					CString strClass;
					BSTR bstrClass;
					hr2 = pADs->get_Class(&bstrClass);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strClass=bstrClass;
#ifdef _DEBUG
	tmp.Format("DEBUGGING: class %s...<br>",strClass);
	pRequest->Write(tmp);
#endif
					
					//
					if(strClass.Compare("IIsWebServer")==0)
					{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: gather webserver info...<br>");
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
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDescription = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AnonymousUserName",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAnonymousUsername = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AnonymousUserPass",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAnonymousPassword = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"ServerBindings",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strBindings,pRequest,';');
						VariantClear(&var);
						
						hr2 = pADs->Get(L"SecureBindings",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strSSLBindings,pRequest,';');
						VariantClear(&var);
						
						hr2 = pADs->Get(L"MaxConnections",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strMaxConnections.Format("%d",V_INT(&var));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"ConnectionTimeout",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strConnectionTimeout.Format("%d",V_INT(&var));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"DefaultLogonDomain",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDefaultLogonDomain = V_BSTR(&var) ;
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
						
						hr2 = pADs->Get(L"EnableDirBrowsing",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDirectoryBrowsingAllowed.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"HTTPErrors",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strCustomErrors,pRequest,';');
						VariantClear(&var);

						// logs
						hr2 = pADs->Get(L"LogType",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strEnableLogging.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);

						hr2 = pADs->Get(L"LogPluginClsid",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2))
						{
							CString str = V_BSTR(&var) ;
							if(str.Compare(logidIIS)==0)
								strActiveLogFormat="IIS";
							else if(str.Compare(logidNCSA)==0)
								strActiveLogFormat="NCSA";
							else if(str.Compare(logidW3C)==0)
								strActiveLogFormat="W3C";
							else if(str.Compare(logidODBC)==0)
								strActiveLogFormat="ODBC";
						}
						VariantClear(&var);

						// ncsa & w3c
						hr2 = pADs->Get(L"LogFileDirectory",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strLogFileDirectory = V_BSTR(&var) ;
						VariantClear(&var);

						hr2 = pADs->Get(L"LogFilePeriod",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2))
						{
							switch(V_INT(&var))
							{
							case 1:
								strNewLogTimePeriod="Daily";
								break;
							case 2:
								strNewLogTimePeriod="Weekly";
								break;
							case 3:
								strNewLogTimePeriod="Monthly";
								break;
							case 4:
								strNewLogTimePeriod="Hourly";
								break;
							case 0:
								strNewLogTimePeriod="Unlimited";
								break;
							}
						}
						VariantClear(&var);

						hr2 = pADs->Get(L"LogFileTruncateSize",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strMB.Format("%d",V_INT(&var));
						VariantClear(&var);

						hr2 = pADs->Get(L"LogFileLocaltimeRollover",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strUseLocalTime.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);

						// log ODBC
						hr2 = pADs->Get(L"LogOdbcDataSource",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strODBCDataSource = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"LogOdbcTableName",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strODBCTable = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"LogOdbcUserName",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strODBCUserName = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"LogOdbcPassword",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strODBCPassword = V_BSTR(&var) ;
						VariantClear(&var);
						

						// Root
						IADs *pADsr = NULL;
						pathe.Format("IIS://%s/W3SVC/%s/Root",computer,strServer);
						hrr=ADsGetObject(A2W(pathe), IID_IADs, (void**) &pADsr );
						//log_COMError(__LINE__,hrr);
#ifdef _DEBUG
	tmp.Format("DEBUGGING: querying subroot container %s...<br>",pathe);
	pRequest->Write(tmp);
#endif

						if(SUCCEEDED(hrr))
						{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: subroot attached...<br>");
#endif
						}

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: handling special cases...<br>");
#endif
						// the Root special cases
						hr2 = pADs->Get(L"HTTPCustomHeaders",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strHTTPHeaders,pRequest,',');
						VariantClear(&var);
						if(strHTTPHeaders.IsEmpty() && SUCCEEDED(hrr))
						{
							hr2 = pADsr->Get(L"HttpCustomHeaders",&var);
							if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strHTTPHeaders,pRequest,',');
							VariantClear(&var);
						}
						
						hr2 = pADs->Get(L"HttpRedirect",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strHTTPRedirect=V_BSTR(&var);
						VariantClear(&var);
						if(strHTTPRedirect.IsEmpty() && SUCCEEDED(hrr))
						{
							hr2 = pADsr->Get(L"HttpRedirect",&var);
							log_COMError(__LINE__,hr2);
							if(SUCCEEDED(hr2)) strHTTPRedirect = V_BSTR(&var) ;
							VariantClear(&var);
						}
						
						hr2 = pADs->Get(L"Path",&var);
						//log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strPath=V_BSTR(&var);
						VariantClear(&var);
						if(strPath.IsEmpty() && SUCCEEDED(hrr))
						{
							hr2 = pADsr->Get(L"Path",&var);
							log_COMError(__LINE__,hr2);
							if(SUCCEEDED(hr2)) strPath = V_BSTR(&var) ;
							VariantClear(&var);
						}
						
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: root cases...<br>");
#endif
						//look to Root
						if(SUCCEEDED(hrr))
						{
						int ap=0;
						hr2 = pADsr->Get(L"AppIsolated",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) ap = V_INT(&var) ;
						VariantClear(&var);
						//ap=3;
						switch (ap)
						{
						case 0: strApplicationProtection="Low"; break;
						case 2: strApplicationProtection="Medium"; break;
						case 1: strApplicationProtection="High"; break;
						default:
							strApplicationProtection="Unknown";
							CString tmp;
							tmp.Format("AppIsolated invalid value of %d (valids 0=in-process, 1=out-of-process, 2=pooled)",ap);
							log_Error(__LINE__,tmp);
						}
						
						hr2 = pADsr->Get(L"AppFriendlyName",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strApplicationName = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADsr->Get(L"AppRoot",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strApplicationStartingPoint = V_BSTR(&var) ;
						VariantClear(&var);
						}
						
						//
						if(SUCCEEDED(hrr))
						{
							hr2=pADsr->Release();
							log_COMError(__LINE__,hr2);
						}

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: creating SQL row...<br>");
#endif
						//
						int iRow = pQuery->AddRow();
						pQuery->SetData( iRow, iServer, strServer );
						pQuery->SetData( iRow, iState, strState );
						pQuery->SetData( iRow, iDescription, strDescription );
						
						pQuery->SetData( iRow, iPath, strPath );
						
						pQuery->SetData( iRow, iAnonymousUsername, strAnonymousUsername );
						pQuery->SetData( iRow, iAnonymousPassword, strAnonymousPassword );
						
						pQuery->SetData( iRow, iBindings, strBindings );
						pQuery->SetData( iRow, iSSLBindings, strSSLBindings );
						
						pQuery->SetData( iRow, iMaxConnections, strMaxConnections );
						pQuery->SetData( iRow, iConnectionTimeout, strConnectionTimeout );
						pQuery->SetData( iRow, iDefaultLogonDomain, strDefaultLogonDomain );
						
						pQuery->SetData( iRow, iEnableDefaultDoc, strEnableDefaultDoc );
						pQuery->SetData( iRow, iDefaultDoc, strDefaultDoc );
						
						pQuery->SetData( iRow, iEnableDocFooter, strEnableDocFooter );
						pQuery->SetData( iRow, iDocFooter, strDocFooter );
						
						pQuery->SetData( iRow, iAccessRead, strAccessRead );
						pQuery->SetData( iRow, iAccessWrite, strAccessWrite );
						pQuery->SetData( iRow, iAccessExecute, strAccessExecute );
						pQuery->SetData( iRow, iAccessScript, strAccessScript );
						pQuery->SetData( iRow, iDirectoryBrowsingAllowed, strDirectoryBrowsingAllowed );
						
						pQuery->SetData( iRow, iCustomErrors, strCustomErrors );
						pQuery->SetData( iRow, iHTTPHeaders, strHTTPHeaders );
						pQuery->SetData( iRow, iHTTPRedirect, strHTTPRedirect );
						
						pQuery->SetData( iRow, iApplicationName, strApplicationName );
						pQuery->SetData( iRow, iApplicationProtection, strApplicationProtection );
						pQuery->SetData( iRow, iApplicationStartingPoint, strApplicationStartingPoint );

						//
						pQuery->SetData( iRow, iEnableLogging, strEnableLogging );
						pQuery->SetData( iRow, iActiveLogFormat, strActiveLogFormat );
						pQuery->SetData( iRow, iLogFileDirectory, strLogFileDirectory );
						pQuery->SetData( iRow, iNewLogTimePeriod, strNewLogTimePeriod );
						pQuery->SetData( iRow, iMB, strMB );
						pQuery->SetData( iRow, iUseLocalTime, strUseLocalTime );
						pQuery->SetData( iRow, iODBCDataSource, strODBCDataSource );
						pQuery->SetData( iRow, iODBCTable, strODBCTable );
						pQuery->SetData( iRow, iODBCUserName, strODBCUserName );
						pQuery->SetData( iRow, iODBCPassword, strODBCPassword );
					}
#ifdef _DEBUG
	pRequest->Write("</ul>DEBUGGING: end enum...<br>");
#endif					
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
	pRequest->Write("DEBUGGING: EXIT: WebServersNT...<br>");
#endif
}


//
//
//
void AddWebServerNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;

#ifdef _DEBUG
	CString tmp;
	pRequest->Write("DEBUGGING: ENTRY: AddWebServerNT<br>");
#endif

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Requireds...<br>");
#endif
	// required
	CString strinDescription = pRequest->GetAttribute( "DESCRIPTION" );
	pRequest->SetVariable( "IISDescription", strinDescription );
	
	CString strinPath = pRequest->GetAttribute( "PATH" );
	pRequest->SetVariable( "IISPath", strinPath );

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Min parms...<br>");
#endif

	// min parms
	if(	strinDescription.IsEmpty() || strinPath.IsEmpty() )
	{
		log_ARGS("AddWebServer","DESCRIPTION and PATH");
		return;
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Optionals...<br>");
#endif
	
	// optional
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	CString strinComputer = pRequest->GetAttribute( "COMPUTER" );
	pRequest->SetVariable( "IISComputer", strinComputer );
	CString computer;
	if(	strinComputer.IsEmpty() )
		computer="LocalHost";
	else
		computer=strinComputer;
	
	CString strinAnonymousUsername = pRequest->GetAttribute( "AnonymousUsername" );
	CString strinAnonymousPassword = pRequest->GetAttribute( "AnonymousPassword" );
	
	CString strinBindings = pRequest->GetAttribute( "Bindings" );
	CString strinSSLBindings = pRequest->GetAttribute( "SSLBindings" );
	
	CString strinMaxConnections = pRequest->GetAttribute( "MaxConnections" );
	CString strinConnectionTimeout = pRequest->GetAttribute( "ConnectionTimeout" );
	CString strinDefaultLogonDomain = pRequest->GetAttribute( "DefaultLogonDomain" );
	
	CString strinEnableDefaultDoc = pRequest->GetAttribute( "EnableDefaultDoc" );
	CString strinDefaultDoc = pRequest->GetAttribute( "DefaultDoc" );
	
	CString strinEnableDocFooter = pRequest->GetAttribute( "EnableDocFooter" );
	CString strinDocFooter = pRequest->GetAttribute( "DocFooter" );
	
	CString strinAccessRead = pRequest->GetAttribute( "AccessRead" );
	CString strinAccessWrite = pRequest->GetAttribute( "AccessWrite" );
	CString strinAccessExecute = pRequest->GetAttribute( "AccessExecute" );
	CString strinAccessScript = pRequest->GetAttribute( "AccessScript" );
	CString strinDirectoryBrowsingAllowed = pRequest->GetAttribute( "DirectoryBrowsingAllowed" );
	
	CString strinCustomErrors = pRequest->GetAttribute( "CustomErrors" );
	CString strinHTTPHeaders = pRequest->GetAttribute( "HTTPHeaders" );
	CString strinHTTPRedirect = pRequest->GetAttribute( "HTTPRedirect" );
	
	CString strinApplicationName = pRequest->GetAttribute( "ApplicationName" );
	CString strinApplicationProtection = pRequest->GetAttribute( "ApplicationProtection" );
	CString strinApplicationStartingPoint = pRequest->GetAttribute( "ApplicationStartingPoint" );

	//
	CString strinEnableLogging = pRequest->GetAttribute( "EnableLogging" );
	CString strinActiveLogFormat = pRequest->GetAttribute( "ActiveLogFormat" );
	CString strinLogFileDirectory = pRequest->GetAttribute( "LogFileDirectory" );
	CString strinNewLogTimePeriod = pRequest->GetAttribute( "NewLogTimePeriod" );
	CString strinMB = pRequest->GetAttribute( "MB" );
	CString strinUseLocalTime = pRequest->GetAttribute( "UseLocalTime" );
	CString strinODBCDataSource = pRequest->GetAttribute( "ODBCDataSource" );
	CString strinODBCTable = pRequest->GetAttribute( "ODBCTable" );
	CString strinODBCUserName = pRequest->GetAttribute( "ODBCUserName" );
	CString strinODBCPassword = pRequest->GetAttribute( "ODBCPassword" );
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Getting next server...<br>");
#endif
	
	CString strServer;
	if(strinServer.IsEmpty())
	{
		int serverno=NextServer(pRequest,computer,"W3SVC");
		strServer.Format("%u",serverno);
	}
	else
		strServer=strinServer;
	pRequest->SetVariable( "IISServer", strServer );
	
	//
	HRESULT hr,hr2;
	CString pathe;

	// get logging module names
	CString logidIIS;
	CString logidW3C;
	CString logidNCSA;
	CString logidODBC;
    IADs *pADs=NULL;
	{
		VARIANT var;
		VariantInit(&var);

		// iis
		pADs=NULL;
		pathe.Format("IIS://%s/logging/Microsoft IIS Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidIIS = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// w3c
		pADs=NULL;
		pathe.Format("IIS://%s/logging/W3C Extended Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidW3C = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// ncsa
		pADs=NULL;
		pathe.Format("IIS://%s/logging/NCSA Common Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidNCSA = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// odbc logging
		pADs=NULL;
		pathe.Format("IIS://%s/logging/ODBC Logging",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidODBC = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
	}



	//
	// PART 1
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: PHASE 1<br>");
#endif

	//
	pathe.Format("IIS://%s/W3SVC",computer);
#ifdef _DEBUG
	tmp.Format("DEBUGGING: iis container %s (computer %s)<br>",pathe,computer);
	pRequest->Write(tmp);
#endif
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Trying to create WebServer object<br>");
#endif
		//
		IDispatch *pDisp=NULL;
		hr = pContainer->Create(L"IIsWebServer", A2W(strServer), &pDisp );
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: WebServer object created<br>");
	pRequest->Write("DEBUGGING: Getting new interface...<br>");
#endif
			//
			IADs *pADs=NULL;
			hr = pDisp->QueryInterface( IID_IADs, (void**)&pADs );
			log_COMError(__LINE__,hr);
			
			//
			hr2=pDisp->Release();
			log_COMError(__LINE__,hr2);
			
			//
			if (SUCCEEDED(hr))
			{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: New object's interface attached...<br>");
#endif
				VARIANT var;
				VariantInit(&var);
				
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Setting required properties...<br>");
#endif
				//set required properties
				V_BSTR(&var) = SysAllocString(A2W(strinDescription));
				V_VT(&var) = VT_BSTR;
				hr2=pADs->Put(L"ServerComment",var);
				log_COMError(__LINE__,hr2);
				VariantClear(&var);
				
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Setting optional properties...<br>");
#endif
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
				
				if(!strinBindings.IsEmpty())
				{
					CStringtoVARIANTARRAY(strinBindings,&var,pRequest,';');
					hr2=pADs->Put(L"ServerBindings",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinSSLBindings.IsEmpty())
				{
					CStringtoVARIANTARRAY(strinSSLBindings,&var,pRequest,';');
					hr2=pADs->Put(L"SecureBindings",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinMaxConnections.IsEmpty())
				{
					V_INT(&var) = atoi(strinMaxConnections);
					V_VT(&var) = VT_INT;
					hr2=pADs->Put(L"MaxConnections",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinConnectionTimeout.IsEmpty())
				{
					V_INT(&var) = atoi(strinConnectionTimeout);
					V_VT(&var) = VT_INT;
					hr2=pADs->Put(L"ConnectionTimeout",var);
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
				
				if(!strinAccessScript.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAccessScript);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AccessScript",var);
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
				
				if(!strinDirectoryBrowsingAllowed.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinDirectoryBrowsingAllowed);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"EnableDirBrowsing",var);
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

				//
				if(!strinEnableLogging.IsEmpty())
				{
					if(strinEnableLogging.Compare("0")==0)
					{
						V_BOOL(&var) = 0;
						V_VT(&var) = VT_BOOL;
						hr2=pADs->Put(L"LogType",var);
						log_COMError(__LINE__,hr2);
						VariantClear(&var);
					}
					else
					{
						V_BOOL(&var) = 1;
						V_VT(&var) = VT_BOOL;
						hr2=pADs->Put(L"LogType",var);
						log_COMError(__LINE__,hr2);
						VariantClear(&var);

						if(strinActiveLogFormat.Compare("IIS")==0)
							V_BSTR(&var) = SysAllocString(A2W(logidIIS));
						else if(strinActiveLogFormat.Compare("NCSA")==0)
							V_BSTR(&var) = SysAllocString(A2W(logidNCSA));
						else if(strinActiveLogFormat.Compare("W3C")==0)
							V_BSTR(&var) = SysAllocString(A2W(logidW3C));
						else if(strinActiveLogFormat.Compare("ODBC")==0)
							V_BSTR(&var) = SysAllocString(A2W(logidODBC));
						else
							V_BSTR(&var) = SysAllocString(L"");
						V_VT(&var) = VT_BSTR;
						hr2=pADs->Put(L"LogPluginClsid",var);
						log_COMError(__LINE__,hr2);
						VariantClear(&var);
					}
				}

				if(!strinLogFileDirectory.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinLogFileDirectory));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"LogFileDirectory",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

				if(!strinNewLogTimePeriod.IsEmpty())
				{
					if(strinNewLogTimePeriod.CompareNoCase("Daily")==0)
						V_INT(&var) = 1;
					else if(strinNewLogTimePeriod.CompareNoCase("Weekly")==0)
						V_INT(&var) = 2;
					else if(strinNewLogTimePeriod.CompareNoCase("Monthly")==0)
						V_INT(&var) = 3;
					else if(strinNewLogTimePeriod.CompareNoCase("Hourly")==0)
						V_INT(&var) = 4;
					else if(strinNewLogTimePeriod.CompareNoCase("Unlimited")==0)
						V_INT(&var) = 0;
					else
						V_INT(&var) = 1;
					V_VT(&var) = VT_INT;
					hr2=pADs->Put(L"LogFilePeriod",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

				if(!strinMB.IsEmpty())
				{
					V_INT(&var) = atoi(strinMB)*1048576;
					V_VT(&var) = VT_INT;
					hr2=pADs->Put(L"LogFileTruncateSize",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

				if(!strinUseLocalTime.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinUseLocalTime);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"LogFileLocaltimeRollover",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

				if(!strinODBCDataSource.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinODBCDataSource));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"LogOdbcDataSource",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

				if(!strinODBCTable.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinODBCTable));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"LogOdbcTableName",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

				if(!strinODBCUserName.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinODBCUserName));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"LogOdbcUserName",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

				if(!strinODBCPassword.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinODBCPassword));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"LogOdbcPassword",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Flushing out changes...<br>");
#endif
				// write it out
				hr2=pADs->SetInfo();
				log_COMError(__LINE__,hr2);
				
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Releasing ADs object interface...<br>");
#endif
				//
				hr2=pADs->Release();
				log_COMError(__LINE__,hr2);
			}
			
			//
			// PART 2
			//
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: PHASE 2<br>");
#endif
			pathe.Format("IIS://%s/W3SVC/%s",computer,strServer);

#ifdef _DEBUG
	tmp.Format("DEBUGGING: iis root container %s (computer %s, server %s)<br>",pathe,computer,strServer);
	pRequest->Write(tmp);
#endif

			IADsContainer *pContainerRoot=NULL;
			hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainerRoot);
			log_COMError(__LINE__,hr);
			
			//
			if(SUCCEEDED(hr))
			{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: PHASE 2 START<br>");
	pRequest->Write("DEBUGGING: Trying to create WebVirtualDir object<br>");
#endif
				//
				IDispatch *pDispRoot=NULL;
				hr = pContainerRoot->Create(L"IIsWebVirtualDir", L"Root", &pDispRoot );
				log_COMError(__LINE__,hr);
				
				//
				hr2=pContainerRoot->Release();
				log_COMError(__LINE__,hr2);
				
				//
				if(SUCCEEDED(hr))
				{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: WebVitualDir object created<br>");
	pRequest->Write("DEBUGGING: Getting new interface...<br>");
#endif
					//
					IADs *pADsRoot=NULL;
					hr = pDispRoot->QueryInterface( IID_IADs, (void**)&pADsRoot );
					log_COMError(__LINE__,hr);
					
					//
					hr2=pDispRoot->Release();
					log_COMError(__LINE__,hr2);
					
					//
					if (SUCCEEDED(hr))
					{
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: New interface attached...<br>");
#endif
						VARIANT var;
						VariantInit(&var);
						
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Setting default properties...<br>");
#endif
						//set properties
						V_BSTR(&var) = SysAllocString(L"IIsWebVirtualDir");
						V_VT(&var) = VT_BSTR;
						hr2=pADsRoot->Put(L"KeyType",var);
						log_COMError(__LINE__,hr2);
						
						V_BSTR(&var) = SysAllocString(A2W(strinPath));
						V_VT(&var) = VT_BSTR;
						hr2=pADsRoot->Put(L"Path",var);
						log_COMError(__LINE__,hr2);
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: setting optional properties...<br>");
#endif
						// optionals
						if(!strinDocFooter.IsEmpty())
						{
							CString tmp;
							tmp="FILE:" + strinDocFooter;
							V_BSTR(&var) = SysAllocString(A2W(tmp));
							V_VT(&var) = VT_BSTR;
							hr2=pADsRoot->Put(L"DefaultDocFooter",var);
							log_COMError(__LINE__,hr2);
							VariantClear(&var);
						}
						
						if(!strinApplicationName.IsEmpty() || !strinApplicationProtection.IsEmpty() )
						{
							V_BSTR(&var) = SysAllocString(A2W(strinApplicationName));
							V_VT(&var) = VT_BSTR;
							hr2=pADsRoot->Put(L"AppFriendlyName",var);
							log_COMError(__LINE__,hr2);
							VariantClear(&var);
							
							int ap=2;
							if(strinApplicationProtection.CompareNoCase("Low")==0)
								ap=0;
							else if(strinApplicationProtection.CompareNoCase("Medium")==0)
								ap=2;
							else if(strinApplicationProtection.CompareNoCase("High")==0)
								ap=1;
							V_INT(&var) = ap;
							V_VT(&var) = VT_INT;
							hr2=pADsRoot->Put(L"AppIsolated",var);
							log_COMError(__LINE__,hr2);
							VariantClear(&var);
							
							CString lm;
							lm.Format("/LM/W3SVC/%s/Root",strServer);
							V_BSTR(&var) = SysAllocString(A2W(lm));
							V_VT(&var) = VT_BSTR;
							hr2=pADsRoot->Put(L"AppRoot",var);
							log_COMError(__LINE__,hr2);
							VariantClear(&var);
						}
						
						if(!strinHTTPHeaders.IsEmpty())
						{
							CStringtoVARIANTARRAY(strinHTTPHeaders,&var,pRequest,',');
							hr2=pADsRoot->Put(L"HttpCustomHeaders",var);
							log_COMError(__LINE__,hr2);
							VariantClear(&var);
						}
						
						if(!strinHTTPRedirect.IsEmpty())
						{
							V_BSTR(&var) = SysAllocString(A2W(strinHTTPRedirect));
							V_VT(&var) = VT_BSTR;
							hr2=pADsRoot->Put(L"HttpRedirect",var);
							log_COMError(__LINE__,hr2);
							VariantClear(&var);
						}
						
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: flushing out changes to AD...<br>");
#endif
						// write it out
						hr2=pADsRoot->SetInfo();
						log_COMError(__LINE__,hr2);
						
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Releasing ADs interface...<br>");
#endif
						//
						hr2=pADsRoot->Release();
						log_COMError(__LINE__,hr2);
					}
				}
			}

		}
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: AddWebServerNT<br>");
#endif
}


//
// required: server
// optional: computer
void DeleteWebServerNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: DeleteWebServerNT<br>");
#endif

	// Required
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	// min parms
	if(	strinServer.IsEmpty() )
	{
		log_ARGS("DeleteWebServer","SERVER");
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
	pathe.Format("IIS://%s/W3SVC",computer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		hr=pContainer->Delete(L"IIsWebServer",A2W(strinServer));
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: DeleteWebServerNT<br>");
#endif
}





//
//
void WebServerNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	CString tmp;
	pRequest->Write("DEBUGGING: ENTRY: WebServerNT<br>");
#endif

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: Establishing SQL fields...<br>");
#endif

	// fields
	CCFXStringSet* pColumns = pRequest->CreateStringSet();
	int iServer = pColumns->AddString( "Server" );
	int iState = pColumns->AddString( "State" );
	int iDescription = pColumns->AddString( "Description" );
	
	int iPath = pColumns->AddString( "Path" );
	
	int iAnonymousUsername = pColumns->AddString( "AnonymousUsername" );
	int iAnonymousPassword = pColumns->AddString( "AnonymousPassword" );
	
	int iBindings = pColumns->AddString( "Bindings" );
	int iSSLBindings = pColumns->AddString( "SSLBindings" );
	
	int iMaxConnections = pColumns->AddString( "MaxConnections" );
	int iConnectionTimeout = pColumns->AddString( "ConnectionTimeout" );
	int iDefaultLogonDomain = pColumns->AddString( "DefaultLogonDomain" );
	
	int iEnableDefaultDoc = pColumns->AddString( "EnableDefaultDoc" );
	int iDefaultDoc = pColumns->AddString( "DefaultDoc" );
	
	int iEnableDocFooter = pColumns->AddString( "EnableDocFooter" );
	int iDocFooter = pColumns->AddString( "DocFooter" );
	
	int iAccessRead = pColumns->AddString( "AccessRead" );
	int iAccessWrite = pColumns->AddString( "AccessWrite" );
	int iAccessExecute = pColumns->AddString( "AccessExecute" );
	int iAccessScript = pColumns->AddString( "AccessScript" );
	int iDirectoryBrowsingAllowed = pColumns->AddString( "DirectoryBrowsingAllowed" );
	
	int iCustomErrors = pColumns->AddString( "CustomErrors" );
	int iHTTPHeaders = pColumns->AddString( "HTTPHeaders" );
	int iHTTPRedirect = pColumns->AddString( "HTTPRedirect" );
	
	int iApplicationName = pColumns->AddString( "ApplicationName" );
	int iApplicationProtection = pColumns->AddString( "ApplicationProtection" );
	int iApplicationStartingPoint = pColumns->AddString( "ApplicationStartingPoint" );

	//
	int iEnableLogging = pColumns->AddString( "EnableLogging" );
	int iActiveLogFormat = pColumns->AddString( "ActiveLogFormat" );
	int iLogFileDirectory = pColumns->AddString( "LogFileDirectory" );
	int iNewLogTimePeriod = pColumns->AddString( "NewLogTimePeriod" );
	int iMB = pColumns->AddString( "MB" );
	int iUseLocalTime = pColumns->AddString( "UseLocalTime" );
	int iODBCDataSource = pColumns->AddString( "ODBCDataSource" );
	int iODBCTable = pColumns->AddString( "ODBCTable" );
	int iODBCUserName = pColumns->AddString( "ODBCUserName" );
	int iODBCPassword = pColumns->AddString( "ODBCPassword" );

	CCFXQuery* pQuery = pRequest->AddQuery( get_query_variable(), pColumns );

	
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

	
	//
    HRESULT hr,hr2,hrr;
	CString pathe;

	
	// get logging module names
	CString logidIIS;
	CString logidW3C;
	CString logidNCSA;
	CString logidODBC;
    IADs *pADs=NULL;
	{
		VARIANT var;
		VariantInit(&var);

		// iis
		pADs=NULL;
		pathe.Format("IIS://%s/logging/Microsoft IIS Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidIIS = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// w3c
		pADs=NULL;
		pathe.Format("IIS://%s/logging/W3C Extended Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidW3C = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// ncsa
		pADs=NULL;
		pathe.Format("IIS://%s/logging/NCSA Common Log File Format",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidNCSA = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
		
		// odbc logging
		pADs=NULL;
		pathe.Format("IIS://%s/logging/ODBC Logging",computer);
		hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
		log_COMError(__LINE__,hr);
		if(SUCCEEDED(hr) && pADs)
		{
			hr2 = pADs->Get(L"LogModuleId",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2))
			{
				logidODBC = V_BSTR(&var) ;
			}
			VariantClear(&var);
		}
	}



	//
	pathe.Format("IIS://%s/W3SVC/%s",computer,strinServer);
#ifdef _DEBUG
	tmp.Format("DEBUGGING: container %s...<br>",pathe);
	pRequest->Write(tmp);
#endif
	
	pADs=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
#ifdef _DEBUG
		pRequest->Write("DEBUGGING: container attached...<br>");
#endif
		
			//
			//VARIANT var;
			
			//
			//			IADs *pADs=(IADs*)var.pdispVal;
			
			//
			//CString strServer;
			CString strState; //ServerState
			CString strDescription; //ServerComment;
			
			CString strPath;
			
			CString strAnonymousUsername;
			CString strAnonymousPassword;
			
			CString strBindings; //ServerBindings
			CString strSSLBindings; //SecureBindings
			
			CString strMaxConnections;
			CString strConnectionTimeout;
			CString strDefaultLogonDomain;
			
			CString strEnableDefaultDoc;
			CString strDefaultDoc;
			
			CString strEnableDocFooter;
			CString strDocFooter;
			
			CString strAccessRead;
			CString strAccessWrite; 
			CString strAccessExecute; 
			CString strAccessScript; 
			CString strDirectoryBrowsingAllowed; //EnableDirBrowsing
			
			CString strCustomErrors; //HTTPErrors
			CString strHTTPHeaders; //HTTPCustomHeaders
			CString strHTTPRedirect;
			
			CString strApplicationName;
			CString strApplicationProtection;
			CString strApplicationStartingPoint;
			
			//
			CString strEnableLogging;
			CString strActiveLogFormat;
			CString strLogFileDirectory;
			CString strNewLogTimePeriod;
			CString strMB;
			CString strUseLocalTime;
			CString strODBCDataSource;
			CString strODBCTable;
			CString strODBCUserName;
			CString strODBCPassword;
			
			//
			/*
			BSTR bstrName;
			hr2 = pADs->get_Name(&bstrName);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strServer=bstrName;
#ifdef _DEBUG
			tmp.Format("DEBUGGING: +server %s...<br>",strServer);
			pRequest->Write(tmp);
#endif
			*/
			//
			CString strClass;
			BSTR bstrClass;
			hr2 = pADs->get_Class(&bstrClass);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strClass=bstrClass;
#ifdef _DEBUG
			tmp.Format("DEBUGGING: class %s...<br>",strClass);
			pRequest->Write(tmp);
#endif
			
			//
			if(strClass.Compare("IIsWebServer")==0)
			{
#ifdef _DEBUG
				pRequest->Write("DEBUGGING: gather webserver info...<br>");
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
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strDescription = V_BSTR(&var) ;
				VariantClear(&var);
				
				hr2 = pADs->Get(L"AnonymousUserName",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strAnonymousUsername = V_BSTR(&var) ;
				VariantClear(&var);
				
				hr2 = pADs->Get(L"AnonymousUserPass",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strAnonymousPassword = V_BSTR(&var) ;
				VariantClear(&var);
				
				hr2 = pADs->Get(L"ServerBindings",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strBindings,pRequest,';');
				VariantClear(&var);
				
				hr2 = pADs->Get(L"SecureBindings",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strSSLBindings,pRequest,';');
				VariantClear(&var);
				
				hr2 = pADs->Get(L"MaxConnections",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strMaxConnections.Format("%d",V_INT(&var));
				VariantClear(&var);
				
				hr2 = pADs->Get(L"ConnectionTimeout",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strConnectionTimeout.Format("%d",V_INT(&var));
				VariantClear(&var);
				
				hr2 = pADs->Get(L"DefaultLogonDomain",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strDefaultLogonDomain = V_BSTR(&var) ;
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
				
				hr2 = pADs->Get(L"EnableDirBrowsing",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strDirectoryBrowsingAllowed.Format("%d",abs(V_BOOL(&var)));
				VariantClear(&var);
				
				hr2 = pADs->Get(L"HTTPErrors",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strCustomErrors,pRequest,';');
				VariantClear(&var);
				
				// logs
				hr2 = pADs->Get(L"LogType",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strEnableLogging.Format("%d",abs(V_BOOL(&var)));
				VariantClear(&var);
				
				hr2 = pADs->Get(L"LogPluginClsid",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2))
				{
					CString str = V_BSTR(&var) ;
					if(str.Compare(logidIIS)==0)
						strActiveLogFormat="IIS";
					else if(str.Compare(logidNCSA)==0)
						strActiveLogFormat="NCSA";
					else if(str.Compare(logidW3C)==0)
						strActiveLogFormat="W3C";
					else if(str.Compare(logidODBC)==0)
						strActiveLogFormat="ODBC";
				}
				VariantClear(&var);
				
				// ncsa & w3c
				hr2 = pADs->Get(L"LogFileDirectory",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strLogFileDirectory = V_BSTR(&var) ;
				VariantClear(&var);
				
				hr2 = pADs->Get(L"LogFilePeriod",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2))
				{
					switch(V_INT(&var))
					{
					case 1:
						strNewLogTimePeriod="Daily";
						break;
					case 2:
						strNewLogTimePeriod="Weekly";
						break;
					case 3:
						strNewLogTimePeriod="Monthly";
						break;
					case 4:
						strNewLogTimePeriod="Hourly";
						break;
					case 0:
						strNewLogTimePeriod="Unlimited";
						break;
					}
				}
				VariantClear(&var);
				
				hr2 = pADs->Get(L"LogFileTruncateSize",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strMB.Format("%d",V_INT(&var));
				VariantClear(&var);
				
				hr2 = pADs->Get(L"LogFileLocaltimeRollover",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strUseLocalTime.Format("%d",abs(V_BOOL(&var)));
				VariantClear(&var);
				
				// log ODBC
				hr2 = pADs->Get(L"LogOdbcDataSource",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strODBCDataSource = V_BSTR(&var) ;
				VariantClear(&var);
				
				hr2 = pADs->Get(L"LogOdbcTableName",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strODBCTable = V_BSTR(&var) ;
				VariantClear(&var);
				
				hr2 = pADs->Get(L"LogOdbcUserName",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strODBCUserName = V_BSTR(&var) ;
				VariantClear(&var);
				
				hr2 = pADs->Get(L"LogOdbcPassword",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strODBCPassword = V_BSTR(&var) ;
				VariantClear(&var);
				
				
				// Root
				IADs *pADsr = NULL;
				pathe.Format("IIS://%s/W3SVC/%s/Root",computer,strinServer);
				hrr=ADsGetObject(A2W(pathe), IID_IADs, (void**) &pADsr );
				//log_COMError(__LINE__,hrr);
#ifdef _DEBUG
				tmp.Format("DEBUGGING: querying subroot container %s...<br>",pathe);
				pRequest->Write(tmp);
#endif
				
				if(SUCCEEDED(hrr))
				{
#ifdef _DEBUG
					pRequest->Write("DEBUGGING: subroot attached...<br>");
#endif
				}
				
#ifdef _DEBUG
				pRequest->Write("DEBUGGING: handling special cases...<br>");
#endif
				// the Root special cases
				hr2 = pADs->Get(L"HTTPCustomHeaders",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strHTTPHeaders,pRequest,',');
				VariantClear(&var);
				if(strHTTPHeaders.IsEmpty() && SUCCEEDED(hrr))
				{
					hr2 = pADsr->Get(L"HttpCustomHeaders",&var);
					if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strHTTPHeaders,pRequest,',');
					VariantClear(&var);
				}
				
				hr2 = pADs->Get(L"HttpRedirect",&var);
				log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strHTTPRedirect=V_BSTR(&var);
				VariantClear(&var);
				if(strHTTPRedirect.IsEmpty() && SUCCEEDED(hrr))
				{
					hr2 = pADsr->Get(L"HttpRedirect",&var);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strHTTPRedirect = V_BSTR(&var) ;
					VariantClear(&var);
				}
				
				hr2 = pADs->Get(L"Path",&var);
				//log_COMError(__LINE__,hr2);
				if(SUCCEEDED(hr2)) strPath=V_BSTR(&var);
				VariantClear(&var);
				if(strPath.IsEmpty() && SUCCEEDED(hrr))
				{
					hr2 = pADsr->Get(L"Path",&var);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strPath = V_BSTR(&var) ;
					VariantClear(&var);
				}
				
#ifdef _DEBUG
				pRequest->Write("DEBUGGING: root cases...<br>");
#endif
				//look to Root
				if(SUCCEEDED(hrr))
				{
					int ap=0;
					hr2 = pADsr->Get(L"AppIsolated",&var);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) ap = V_INT(&var) ;
					VariantClear(&var);
					//ap=3;
					switch (ap)
					{
					case 0: strApplicationProtection="Low"; break;
					case 2: strApplicationProtection="Medium"; break;
					case 1: strApplicationProtection="High"; break;
					default:
						strApplicationProtection="Unknown";
						CString tmp;
						tmp.Format("AppIsolated invalid value of %d (valids 0=in-process, 1=out-of-process, 2=pooled)",ap);
						log_Error(__LINE__,tmp);
					}
					
					hr2 = pADsr->Get(L"AppFriendlyName",&var);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strApplicationName = V_BSTR(&var) ;
					VariantClear(&var);
					
					hr2 = pADsr->Get(L"AppRoot",&var);
					log_COMError(__LINE__,hr2);
					if(SUCCEEDED(hr2)) strApplicationStartingPoint = V_BSTR(&var) ;
					VariantClear(&var);
				}
				
				//
				if(SUCCEEDED(hrr))
				{
					hr2=pADsr->Release();
					log_COMError(__LINE__,hr2);
				}
				
#ifdef _DEBUG
				pRequest->Write("DEBUGGING: creating SQL row...<br>");
#endif
				//
				int iRow = pQuery->AddRow();
				pQuery->SetData( iRow, iServer, strinServer );
				pQuery->SetData( iRow, iState, strState );
				pQuery->SetData( iRow, iDescription, strDescription );
				
				pQuery->SetData( iRow, iPath, strPath );
				
				pQuery->SetData( iRow, iAnonymousUsername, strAnonymousUsername );
				pQuery->SetData( iRow, iAnonymousPassword, strAnonymousPassword );
				
				pQuery->SetData( iRow, iBindings, strBindings );
				pQuery->SetData( iRow, iSSLBindings, strSSLBindings );
				
				pQuery->SetData( iRow, iMaxConnections, strMaxConnections );
				pQuery->SetData( iRow, iConnectionTimeout, strConnectionTimeout );
				pQuery->SetData( iRow, iDefaultLogonDomain, strDefaultLogonDomain );
				
				pQuery->SetData( iRow, iEnableDefaultDoc, strEnableDefaultDoc );
				pQuery->SetData( iRow, iDefaultDoc, strDefaultDoc );
				
				pQuery->SetData( iRow, iEnableDocFooter, strEnableDocFooter );
				pQuery->SetData( iRow, iDocFooter, strDocFooter );
				
				pQuery->SetData( iRow, iAccessRead, strAccessRead );
				pQuery->SetData( iRow, iAccessWrite, strAccessWrite );
				pQuery->SetData( iRow, iAccessExecute, strAccessExecute );
				pQuery->SetData( iRow, iAccessScript, strAccessScript );
				pQuery->SetData( iRow, iDirectoryBrowsingAllowed, strDirectoryBrowsingAllowed );
				
				pQuery->SetData( iRow, iCustomErrors, strCustomErrors );
				pQuery->SetData( iRow, iHTTPHeaders, strHTTPHeaders );
				pQuery->SetData( iRow, iHTTPRedirect, strHTTPRedirect );
				
				pQuery->SetData( iRow, iApplicationName, strApplicationName );
				pQuery->SetData( iRow, iApplicationProtection, strApplicationProtection );
				pQuery->SetData( iRow, iApplicationStartingPoint, strApplicationStartingPoint );
				
				//
				pQuery->SetData( iRow, iEnableLogging, strEnableLogging );
				pQuery->SetData( iRow, iActiveLogFormat, strActiveLogFormat );
				pQuery->SetData( iRow, iLogFileDirectory, strLogFileDirectory );
				pQuery->SetData( iRow, iNewLogTimePeriod, strNewLogTimePeriod );
				pQuery->SetData( iRow, iMB, strMB );
				pQuery->SetData( iRow, iUseLocalTime, strUseLocalTime );
				pQuery->SetData( iRow, iODBCDataSource, strODBCDataSource );
				pQuery->SetData( iRow, iODBCTable, strODBCTable );
				pQuery->SetData( iRow, iODBCUserName, strODBCUserName );
				pQuery->SetData( iRow, iODBCPassword, strODBCPassword );
			}
#ifdef _DEBUG
			pRequest->Write("</ul>DEBUGGING: end enum...<br>");
#endif					
			//
			hr2=pADs->Release();
			log_COMError(__LINE__,hr2);
					
			//
			SysFreeString(bstrClass);
//			SysFreeString(bstrName);
		}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: WebServerNT...<br>");
#endif
}

// old trick to get the total number of lines of code
int getlines_webserver() { return __LINE__+1; }
