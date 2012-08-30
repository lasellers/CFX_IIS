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

#include "ftp.h"

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

// ftp server

//
// required:
// optional: computer
void FTPServersNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: FTPServersNT<br>");
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
	
	int iPath = pColumns->AddString( "Path" );
	
	int iAnonymousUsername = pColumns->AddString( "AnonymousUsername" );
	int iAnonymousPassword = pColumns->AddString( "AnonymousPassword" );
	int iAnonymousPasswordSync = pColumns->AddString( "AnonymousPasswordSync" );
	
	int iBindings = pColumns->AddString( "Bindings" );
	
	int iMaxConnections = pColumns->AddString( "MaxConnections" );
	int iConnectionTimeout = pColumns->AddString( "ConnectionTimeout" );
	int iDefaultLogonDomain = pColumns->AddString( "DefaultLogonDomain" );
	
	int iAccessRead = pColumns->AddString( "AccessRead" );
	int iAccessWrite = pColumns->AddString( "AccessWrite" );
	int iLogAccess = pColumns->AddString( "LogAccess" );
	int iDirectoryBrowsingAllowed = pColumns->AddString( "DirectoryBrowsingAllowed" );
	
	int iExitMessage = pColumns->AddString( "ExitMessage" );
	int iGreetingMessage = pColumns->AddString( "GreetingMessage" );
	int iMaxClientsMessage = pColumns->AddString( "MaxClientsMessage" );
	
	int iAnonymousOnly = pColumns->AddString( "AnonymousOnly" );
	int iAllowAnonymous = pColumns->AddString( "AllowAnonymous" );

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
	pathe.Format("IIS://%s/MSFTPSVC",computer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr) && pContainer)
	{
		//
		IEnumVARIANT *pEnum;
		hr = ADsBuildEnumerator(pContainer, &pEnum); 
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
			//
			VARIANT enumvar;
			ULONG ulElements=1;
			
			//
			while ( (SUCCEEDED(hr)) && (ulElements==1) )
			{
				//
				hr = ADsEnumerateNext(pEnum, 1, &enumvar, &ulElements);
				log_COMError(__LINE__,hr);
				
				//
				if(SUCCEEDED(hr) && (ulElements==1) )
				{
					//
					CString strServer;
					CString strState; //ServerState
					CString strDescription; //ServerComment;
					
					CString strPath;
					
					CString strAnonymousUsername;
					CString strAnonymousPassword;
					CString strAnonymousPasswordSync;
					
					CString strBindings; //ServerBindings
					
					CString strMaxConnections;
					CString strConnectionTimeout;
					CString strDefaultLogonDomain;
					
					CString strAccessRead;
					CString strAccessWrite; 
					CString strLogAccess; 
					CString strDirectoryBrowsingAllowed; //EnableDirBrowsing
					
					CString strExitMessage;
					CString strGreetingMessage;
					CString strMaxClientsMessage;
					
					CString strAnonymousOnly;
					CString strAllowAnonymous;

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
					IADs *pADs=(IADs*)enumvar.pdispVal;
					
					//
					BSTR bstrName;
					hr2 = pADs->get_Name(&bstrName);
					if(SUCCEEDED(hr2)) strServer = bstrName;
					
					//
					BSTR bstrClass;
					hr2 = pADs->get_Class(&bstrClass);
					CString strClass;
					if(SUCCEEDED(hr2)) strClass = bstrClass;
					
					//
					if(strClass.Compare("IIsFtpServer")==0)
					{
						VARIANT var;
						VariantInit(&var);
						
						//
						hr2 = pADs->Get(L"ServerState",&var);
						log_COMError(__LINE__,hr2);
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
						
						hr2 = pADs->Get(L"AnonymousPasswordSync",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAnonymousPasswordSync.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"ServerBindings",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strBindings,pRequest,';');
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
						
						hr2 = pADs->Get(L"AccessRead",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessRead.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AccessWrite",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessWrite.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"DontLog",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strLogAccess.Format("%d",abs(V_BOOL(&var))?0:1);
						VariantClear(&var);
						
						hr2 = pADs->Get(L"EnableDirBrowsing",&var);
						//log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strDirectoryBrowsingAllowed.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"ExitMessage",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strExitMessage = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"GreetingMessage",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strGreetingMessage,pRequest,';');
						VariantClear(&var);
						
						hr2 = pADs->Get(L"MaxClientsMessage",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strMaxClientsMessage = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AnonymousOnly",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAnonymousOnly.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AllowAnonymous",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAllowAnonymous.Format("%d",abs(V_BOOL(&var)));
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

						
						//special-case get path which is down in Root
						HRESULT hrr=S_OK;
						IADsContainer *pContainerr=NULL;
						IADs *pADsr = NULL;
						CString patheroot;
						patheroot.Format("IIS://%s/MSFTPSVC/%s/Root",computer,strServer);
						hrr=ADsGetObject(A2W(patheroot), IID_IADs, (void**) &pADsr );
						log_COMError(__LINE__,hrr);
						if(SUCCEEDED(hrr))
						{
							hr2 = pADsr->Get(L"Path",&var);
							if(SUCCEEDED(hr2)) strPath = V_BSTR(&var) ;
							VariantClear(&var);
						}
						
						//
						int iRow = pQuery->AddRow();
						pQuery->SetData( iRow, iServer, strServer );
						pQuery->SetData( iRow, iState, strState );
						pQuery->SetData( iRow, iDescription, strDescription );
						
						pQuery->SetData( iRow, iPath, strPath );
						
						pQuery->SetData( iRow, iAnonymousUsername, strAnonymousUsername );
						pQuery->SetData( iRow, iAnonymousPassword, strAnonymousPassword );
						pQuery->SetData( iRow, iAnonymousPasswordSync, strAnonymousPasswordSync );
						
						pQuery->SetData( iRow, iBindings, strBindings );
						
						pQuery->SetData( iRow, iMaxConnections, strMaxConnections );
						pQuery->SetData( iRow, iConnectionTimeout, strConnectionTimeout );
						pQuery->SetData( iRow, iDefaultLogonDomain, strDefaultLogonDomain );
						
						pQuery->SetData( iRow, iAccessRead, strAccessRead );
						pQuery->SetData( iRow, iAccessWrite, strAccessWrite );
						pQuery->SetData( iRow, iLogAccess, strLogAccess );
						pQuery->SetData( iRow, iDirectoryBrowsingAllowed, strDirectoryBrowsingAllowed );
						
						pQuery->SetData( iRow, iExitMessage, strExitMessage );
						pQuery->SetData( iRow, iGreetingMessage, strGreetingMessage );
						pQuery->SetData( iRow, iMaxClientsMessage, strMaxClientsMessage );
						
						pQuery->SetData( iRow, iAnonymousOnly, strAnonymousOnly );
						pQuery->SetData( iRow, iAllowAnonymous, strAllowAnonymous );

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
					
					//
					hr2=pADs->Release();
					log_COMError(__LINE__,hr2);
					
					//
					SysFreeString(bstrName);
					SysFreeString(bstrClass);
				}
			}
		}
		
		//		
		hr2 = ADsFreeEnumerator(pEnum);
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: FTPServersNT<br>");
#endif
}





//
// Required: description path
// Optional: computer
void AddFTPServerNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: AddFTPServerNT<br>");
#endif

	// required
	CString strinDescription = pRequest->GetAttribute( "DESCRIPTION" );
	pRequest->SetVariable( "IISDescription", strinDescription );
	
	CString strinPath = pRequest->GetAttribute( "PATH" );
	pRequest->SetVariable( "IISPath", strinPath );
	
	// min parms
	if(	strinDescription.IsEmpty() || strinPath.IsEmpty() )
	{
		log_ARGS("AddFTPServer","DESCRIPTION and PATH");
		return;
	}
	
	// optional
	CString strinServer = pRequest->GetAttribute( "Server" );
	
	CString strinComputer = pRequest->GetAttribute( "COMPUTER" );
	pRequest->SetVariable( "IISComputer", strinComputer );
	CString computer;
	if(	strinComputer.IsEmpty() )
		computer="LocalHost";
	else
		computer=strinComputer;
	
	CString strinAnonymousUsername = pRequest->GetAttribute( "AnonymousUsername" );
	CString strinAnonymousPassword = pRequest->GetAttribute( "AnonymousPassword" );
	CString strinAnonymousPasswordSync = pRequest->GetAttribute( "AnonymousPasswordSync" );
	
	CString strinBindings = pRequest->GetAttribute( "Bindings" );
	
	CString strinMaxConnections = pRequest->GetAttribute( "MaxConnections" );
	CString strinConnectionTimeout = pRequest->GetAttribute( "ConnectionTimeout" );
	CString strinDefaultLogonDomain = pRequest->GetAttribute( "DefaultLogonDomain" );
	
	CString strinAccessRead = pRequest->GetAttribute( "AccessRead" );
	CString strinAccessWrite = pRequest->GetAttribute( "AccessWrite" );
	CString strinLogAccess = pRequest->GetAttribute( "LogAccess" );
	CString strinDirectoryBrowsingAllowed = pRequest->GetAttribute( "DirectoryBrowsing" );
	
	CString strinExitMessage = pRequest->GetAttribute( "ExitMessage" );
	CString strinGreetingMessage = pRequest->GetAttribute( "GreetingMessage" );
	CString strinMaxClientsMessage = pRequest->GetAttribute( "MaxClientMessage" );
	
	CString strinAnonymousOnly = pRequest->GetAttribute( "AnonymousOnly" );
	CString strinAllowAnonymous = pRequest->GetAttribute( "AllowAnonymous" );

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
	
	CString strServer;
	if(strinServer.IsEmpty())
	{
		int serverno=NextServer(pRequest,computer,"MSFTPSVC");
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
	
	//
	pathe.Format("IIS://%s/MSFTPSVC",computer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		// create the ftp virtual server object
		IDispatch *pDisp=NULL;
		hr = pContainer->Create(L"IIsFtpServer", A2W(strServer), &pDisp );
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
			//
			IADs *pADs=NULL;
			hr = pDisp->QueryInterface( IID_IADs, (void**)&pADs ); //get ads interface
			log_COMError(__LINE__,hr);
			
			//
			hr2=pDisp->Release();
			log_COMError(__LINE__,hr2);
			
			//
			if (SUCCEEDED(hr))
			{
				VARIANT var;
				VariantInit(&var);
				//?keytype
				//set required properties
				V_BSTR(&var) = SysAllocString(A2W(strinDescription));
				V_VT(&var) = VT_BSTR;
				hr=pADs->Put(L"ServerComment",var);
				VariantClear(&var);
				
				// optional
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
					hr2=pADs->Put(L"AnonymousPassword",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAnonymousPasswordSync.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinAnonymousPasswordSync));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"AnonymousPasswordSync",var);
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
				
				if(!strinLogAccess.IsEmpty())
				{
					V_BOOL(&var) = abs(atoi(strinLogAccess))?0:1;
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"DontLog",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinExitMessage.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinExitMessage));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"ExitMessage",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinGreetingMessage.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinGreetingMessage));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"GreetingMessage",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinMaxClientsMessage.IsEmpty())
				{
					V_BSTR(&var) = SysAllocString(A2W(strinMaxClientsMessage));
					V_VT(&var) = VT_BSTR;
					hr2=pADs->Put(L"MaxClientsMessage",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAnonymousOnly.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAnonymousOnly);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AnonymousOnly",var);
					log_COMError(__LINE__,hr2);
					VariantClear(&var);
				}
				
				if(!strinAllowAnonymous.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinAllowAnonymous);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"AllowAnonymous",var);
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

				// write it out
				hr2=pADs->SetInfo();
				log_COMError(__LINE__,hr2);
				
				//
				hr2=pADs->Release();
				log_COMError(__LINE__,hr2);
			}
		}
	}
	
	
	//
	// PART 2
	
	//
	pathe.Format("IIS://%s/MSFTPSVC/%s",computer,strServer);
	
	IADsContainer *pContainerRoot=NULL;
	hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainerRoot);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		// create the ftp virtual directory object
		IDispatch *pDispRoot=NULL;
		hr = pContainerRoot->Create(L"IIsFtpVirtualDir", L"Root", &pDispRoot );
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainerRoot->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
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
				VARIANT var;
				VariantInit(&var);
				
				//set properties
				V_BSTR(&var) = SysAllocString(L"IIsFtpVirtualDir");
				V_VT(&var) = VT_BSTR;
				hr2=pADsRoot->Put(L"KeyType",var);
				log_COMError(__LINE__,hr2);
				
				V_BSTR(&var) = SysAllocString(A2W(strinPath));
				V_VT(&var) = VT_BSTR;
				hr2=pADsRoot->Put(L"Path",var);
				log_COMError(__LINE__,hr2);
				
				// write it out
				hr2=pADsRoot->SetInfo();
				log_COMError(__LINE__,hr2);
				
				//
				hr2=pADsRoot->Release();
				log_COMError(__LINE__,hr2);
			}
		}
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: AddFTPServerNT<br>");
#endif
}


//
// required: server
// optional:
void DeleteFTPServerNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: DeleteFTPServerNT<br>");
#endif
	
	// required
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	// min parms
	if(	strinServer.IsEmpty()  )
	{
		log_ARGS("DeleteFTPServer","SERVER");
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
	pathe.Format("IIS://%s/MSFTPSVC",computer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		hr=pContainer->Delete(L"IIsFtpServer",A2W(strinServer));
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: DeleteFTPServerNT<br>");
#endif
}



//
// required:
// optional: computer
void FTPServerNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: FTPServerNT<br>");
#endif

	// required
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	// min parms
	if(	strinServer.IsEmpty() )
	{
		log_ARGS("FTPServer","SERVER");
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
	int iState = pColumns->AddString( "State" );
	int iDescription = pColumns->AddString( "Description" );
	
	int iPath = pColumns->AddString( "Path" );
	
	int iAnonymousUsername = pColumns->AddString( "AnonymousUsername" );
	int iAnonymousPassword = pColumns->AddString( "AnonymousPassword" );
	int iAnonymousPasswordSync = pColumns->AddString( "AnonymousPasswordSync" );
	
	int iBindings = pColumns->AddString( "Bindings" );
	
	int iMaxConnections = pColumns->AddString( "MaxConnections" );
	int iConnectionTimeout = pColumns->AddString( "ConnectionTimeout" );
	int iDefaultLogonDomain = pColumns->AddString( "DefaultLogonDomain" );
	
	int iAccessRead = pColumns->AddString( "AccessRead" );
	int iAccessWrite = pColumns->AddString( "AccessWrite" );
	int iLogAccess = pColumns->AddString( "LogAccess" );
	int iDirectoryBrowsingAllowed = pColumns->AddString( "DirectoryBrowsingAllowed" );
	
	int iExitMessage = pColumns->AddString( "ExitMessage" );
	int iGreetingMessage = pColumns->AddString( "GreetingMessage" );
	int iMaxClientsMessage = pColumns->AddString( "MaxClientsMessage" );
	
	int iAnonymousOnly = pColumns->AddString( "AnonymousOnly" );
	int iAllowAnonymous = pColumns->AddString( "AllowAnonymous" );

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
	pathe.Format("IIS://%s/MSFTPSVC/%s",computer,strinServer);
	
    pADs=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADs, (void**)&pADs);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		CString strState; //ServerState
		CString strDescription; //ServerComment;
		
		CString strPath;
		
		CString strAnonymousUsername;
		CString strAnonymousPassword;
		CString strAnonymousPasswordSync;
		
		CString strBindings; //ServerBindings
		
		CString strMaxConnections;
		CString strConnectionTimeout;
		CString strDefaultLogonDomain;
		
		CString strAccessRead;
		CString strAccessWrite; 
		CString strLogAccess; 
		CString strDirectoryBrowsingAllowed; //EnableDirBrowsing
		
		CString strExitMessage;
		CString strGreetingMessage;
		CString strMaxClientsMessage;
		
		CString strAnonymousOnly;
		CString strAllowAnonymous;
		
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
		BSTR bstrClass;
		hr2 = pADs->get_Class(&bstrClass);
		CString strClass;
		if(SUCCEEDED(hr2)) strClass = bstrClass;
		
		//
		if(strClass.Compare("IIsFtpServer")==0)
		{
			VARIANT var;
			VariantInit(&var);
			
			//
			hr2 = pADs->Get(L"ServerState",&var);
			log_COMError(__LINE__,hr2);
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
			
			hr2 = pADs->Get(L"AnonymousPasswordSync",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strAnonymousPasswordSync.Format("%d",abs(V_BOOL(&var)));
			VariantClear(&var);
			
			hr2 = pADs->Get(L"ServerBindings",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strBindings,pRequest,';');
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
			
			hr2 = pADs->Get(L"AccessRead",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strAccessRead.Format("%d",abs(V_BOOL(&var)));
			VariantClear(&var);
			
			hr2 = pADs->Get(L"AccessWrite",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strAccessWrite.Format("%d",abs(V_BOOL(&var)));
			VariantClear(&var);
			
			hr2 = pADs->Get(L"DontLog",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strLogAccess.Format("%d",abs(V_BOOL(&var))?0:1);
			VariantClear(&var);
			
			hr2 = pADs->Get(L"EnableDirBrowsing",&var);
			//log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strDirectoryBrowsingAllowed.Format("%d",abs(V_BOOL(&var)));
			VariantClear(&var);
			
			hr2 = pADs->Get(L"ExitMessage",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strExitMessage = V_BSTR(&var) ;
			VariantClear(&var);
			
			hr2 = pADs->Get(L"GreetingMessage",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) VARIANTARRAYtoCString(var,strGreetingMessage,pRequest,';');
			VariantClear(&var);
			
			hr2 = pADs->Get(L"MaxClientsMessage",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strMaxClientsMessage = V_BSTR(&var) ;
			VariantClear(&var);
			
			hr2 = pADs->Get(L"AnonymousOnly",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strAnonymousOnly.Format("%d",abs(V_BOOL(&var)));
			VariantClear(&var);
			
			hr2 = pADs->Get(L"AllowAnonymous",&var);
			log_COMError(__LINE__,hr2);
			if(SUCCEEDED(hr2)) strAllowAnonymous.Format("%d",abs(V_BOOL(&var)));
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
			
			
			//special-case get path which is down in Root
			HRESULT hrr=S_OK;
			IADsContainer *pContainerr=NULL;
			IADs *pADsr = NULL;
			CString patheroot;
			patheroot.Format("IIS://%s/MSFTPSVC/%s/Root",computer,strinServer);
			hrr=ADsGetObject(A2W(patheroot), IID_IADs, (void**) &pADsr );
			log_COMError(__LINE__,hrr);
			if(SUCCEEDED(hrr))
			{
				hr2 = pADsr->Get(L"Path",&var);
				if(SUCCEEDED(hr2)) strPath = V_BSTR(&var) ;
				VariantClear(&var);
			}
			
			//
			int iRow = pQuery->AddRow();
			pQuery->SetData( iRow, iServer, strinServer );
			pQuery->SetData( iRow, iState, strState );
			pQuery->SetData( iRow, iDescription, strDescription );
			
			pQuery->SetData( iRow, iPath, strPath );
			
			pQuery->SetData( iRow, iAnonymousUsername, strAnonymousUsername );
			pQuery->SetData( iRow, iAnonymousPassword, strAnonymousPassword );
			pQuery->SetData( iRow, iAnonymousPasswordSync, strAnonymousPasswordSync );
			
			pQuery->SetData( iRow, iBindings, strBindings );
			
			pQuery->SetData( iRow, iMaxConnections, strMaxConnections );
			pQuery->SetData( iRow, iConnectionTimeout, strConnectionTimeout );
			pQuery->SetData( iRow, iDefaultLogonDomain, strDefaultLogonDomain );
			
			pQuery->SetData( iRow, iAccessRead, strAccessRead );
			pQuery->SetData( iRow, iAccessWrite, strAccessWrite );
			pQuery->SetData( iRow, iLogAccess, strLogAccess );
			pQuery->SetData( iRow, iDirectoryBrowsingAllowed, strDirectoryBrowsingAllowed );
			
			pQuery->SetData( iRow, iExitMessage, strExitMessage );
			pQuery->SetData( iRow, iGreetingMessage, strGreetingMessage );
			pQuery->SetData( iRow, iMaxClientsMessage, strMaxClientsMessage );
			
			pQuery->SetData( iRow, iAnonymousOnly, strAnonymousOnly );
			pQuery->SetData( iRow, iAllowAnonymous, strAllowAnonymous );
			
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
					
		//
		hr2=pADs->Release();
		log_COMError(__LINE__,hr2);
					
		//
		SysFreeString(bstrClass);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: FTPServerNT<br>");
#endif
}









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

// ftp directory

//
// Required: SERVER
// optional: computer
void FTPDirectoriesNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: FTPDirectoriesNT<br>");
#endif

	// required
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	// min parms
	if(	strinServer.IsEmpty() )
	{
		log_ARGS("FTPDirectories","SERVER");
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
	int iName = pColumns->AddString( "Name" );
	int iPath = pColumns->AddString( "Path" );
	int iAccessRead = pColumns->AddString( "AccessRead" );
	int iAccessWrite = pColumns->AddString( "AccessWrite" );
	int iLogAccess = pColumns->AddString( "LogAccess" );
	int iUNCUsername = pColumns->AddString( "UNCUsername" );
	int iUNCPassword = pColumns->AddString( "UNCPassword" );
	CCFXQuery* pQuery = pRequest->AddQuery( get_query_variable(), pColumns );
	
	//
	HRESULT hr,hr2;
	CString pathe;
	
	//
	pathe.Format("IIS://%s/MSFTPSVC/%s/Root",computer,strinServer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		IEnumVARIANT *pEnum;
		hr = ADsBuildEnumerator(pContainer, &pEnum); 
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
		
		//
		if(SUCCEEDED(hr))
		{
			//
			VARIANT var;
			ULONG ulElements=1;
			
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
					CString strServer;
					CString strName;
					CString strPath;
					CString strAccessRead;
					CString strAccessWrite; 
					CString strLogAccess; 
					CString strUNCUsername;
					CString strUNCPassword;
					
					//
					IADs *pADs=(IADs*)var.pdispVal;
					
					//
					strServer=strinServer;
					
					BSTR bstrName;
					hr2 = pADs->get_Name(&bstrName);
					if(SUCCEEDED(hr2)) strName=bstrName;
					
					CString strClass;
					BSTR bstrClass;
					hr2 = pADs->get_Class(&bstrClass);
					if(SUCCEEDED(hr2)) strClass=bstrClass;
					
					//
					if(strClass.Compare("IIsFtpVirtualDir")==0)
					{
						VARIANT var;
						VariantInit(&var);
						
						//
						hr2 = pADs->Get(L"ServerComment",&var);
						//log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strName = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"Path",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strPath = V_BSTR(&var) ;
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AccessRead",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessRead.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"AccessWrite",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strAccessWrite.Format("%d",abs(V_BOOL(&var)));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"DontLog",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strLogAccess.Format("%d",abs(V_BOOL(&var))?0:1);
						VariantClear(&var);
						
						hr2 = pADs->Get(L"UNCUserName",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strUNCUsername.Format("%S",V_BSTR(&var));
						VariantClear(&var);
						
						hr2 = pADs->Get(L"UNCPassword",&var);
						log_COMError(__LINE__,hr2);
						if(SUCCEEDED(hr2)) strUNCPassword.Format("%S",V_BSTR(&var));
						VariantClear(&var);
						
						//
						int iRow = pQuery->AddRow();
						pQuery->SetData( iRow, iServer, strServer );
						pQuery->SetData( iRow, iName, strName );
						pQuery->SetData( iRow, iAccessRead, strAccessRead );
						pQuery->SetData( iRow, iAccessWrite, strAccessWrite );
						pQuery->SetData( iRow, iLogAccess, strLogAccess );
						pQuery->SetData( iRow, iPath, strPath );
						pQuery->SetData( iRow, iUNCUsername, strUNCUsername );
						pQuery->SetData( iRow, iUNCPassword, strUNCPassword );
					}
					
					//
					hr2=pADs->Release();
					log_COMError(__LINE__,hr2);
					
					//
					SysFreeString(bstrName);
					SysFreeString(bstrClass);
				}
			}
		}
		//		
		hr2 = ADsFreeEnumerator(pEnum);
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: FTPDirectoriesNT<br>");
#endif
}



//
// Required: server directory path
// Optional: computer uncusername uncpassword
void AddFTPDirectoryNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: AddFTPDirectoryNT<br>");
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
		log_ARGS("AddFTPDirectory","SERVER, DIRECTORY and PATH");
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
	
	CString strinAccessRead = pRequest->GetAttribute( "AccessRead" );
	CString strinAccessWrite = pRequest->GetAttribute( "AccessWrite" );
	CString strinLogAccess = pRequest->GetAttribute( "LogAccess" );
	CString strinUNCUsername = pRequest->GetAttribute( "UNCUsername" );
	CString strinUNCPassword = pRequest->GetAttribute( "UNCPassword" );
	
	//
    HRESULT hr,hr2;
	CString pathe;
	
	//
	pathe.Format("IIS://%s/MSFTPSVC/%s/Root",computer,strinServer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		IDispatch *pDisp=NULL;
		hr = pContainer->Create(L"IIsFtpVirtualDir", A2W(strinDirectory), &pDisp );
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
			hr2=pDisp->Release();
			log_COMError(__LINE__,hr2);
			
			//
			hr2=pADs->SetInfo();
			log_COMError(__LINE__,hr2);
			
			//
			if ( SUCCEEDED(hr) )
			{
				//
				VARIANT var;
				VariantInit(&var);
				
				//set required properties
				V_BSTR(&var) = SysAllocString(L"IIsFtpVirtualDir");
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
				
				if(!strinLogAccess.IsEmpty())
				{
					V_BOOL(&var) = atoi(strinLogAccess);
					V_VT(&var) = VT_BOOL;
					hr2=pADs->Put(L"DontLog",var);
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
				
				// commit changes
				hr2=pADs->SetInfo();
				log_COMError(__LINE__,hr2);
				
				//
				hr2=pADs->Release();
				log_COMError(__LINE__,hr2);
			}
		}
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: AddTPDirectoryNT<br>");
#endif
}



//
// required: server directory
// optional: computer
void DeleteFTPDirectoryNT( CCFXRequest* pRequest )
{
	USES_CONVERSION;
	
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: ENTRY: DeleteFTPDirectoryNT<br>");
#endif

	// required
	CString strinServer = pRequest->GetAttribute( "SERVER" );
	pRequest->SetVariable( "IISServer", strinServer );
	
	CString strinDirectory = pRequest->GetAttribute( "DIRECTORY" );
	pRequest->SetVariable( "IISDirectory", strinDirectory );
	
	// min parms
	if(	strinServer.IsEmpty() || strinDirectory.IsEmpty() )
	{
		log_ARGS("DeleteFTPDirectory","SERVER and DIRECTORY");
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
	pathe.Format("IIS://%s/MSFTPSVC/%s/Root",computer,strinServer);
	
    IADsContainer *pContainer=NULL;
    hr=ADsGetObject(A2W(pathe), IID_IADsContainer, (void**)&pContainer);
	log_COMError(__LINE__,hr);
	
	//
	if(SUCCEEDED(hr))
	{
		//
		hr=pContainer->Delete(L"IIsFtpVirtualDir",A2W(strinDirectory));
		log_COMError(__LINE__,hr);
		
		//
		hr2=pContainer->Release();
		log_COMError(__LINE__,hr2);
	}
#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: DeleteFTPDirectoryNT<br>");
#endif
}






// old trick to get the total number of lines of code
int getlines_ftp() { return __LINE__+1; }
