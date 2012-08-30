/*
IIS7 module management.

http://learn.iis.net/page.aspx/135/discover-installed-components/
http://forums.iis.net/t/1171372.aspx

*/

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

#include "modules.h"

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

// modules

//
/*

HKEY_LOCAL_MACHINE\Software\Microsoft\InetStp\Components\  
Display Name
Registry Key
Web Server
W3SVC
Common HTTP Features
Static Content
StaticContent
Default Document
DefaultDocument
Directory Browsing
DirectoryBrowse
HTTP Errors
HttpErrors
HTTP Redirection
HttpRedirect
Application Development Features 
ASP.NET
ASPNET
.NET Extensibility
NetFxExtensibility
ASP
ASP
CGI
CGI
ISAPI Extensions
ISAPIExtensions
ISAPI Filters
ISAPIFilter
Server-Side Includes
ServerSideInclude
Health and Diagnostics 
HTTP Logging
HttpLogging
Logging Tools
LoggingLibraries
Request Monitor
RequestMonitor
Tracing
HttpTracing
Custom Logging
CustomLogging
ODBC Logging
ODBCLogging
Security
Basic Authentication
BasicAuthentication
Windows Authentication
WindowsAuthentication
Digest Authentication
DigestAuthentication
Client Certificate Mapping Authentication
ClientCertificateMappingAuthentication
IIS Client Certificate Mapping Authentication
IISClientCertificateMappingAuthentication
URL Authorization
Authorization
Request Filtering
RequestFiltering
IP and Domain Restrictions
IPSecurity
Performance Features 
Static Content Compression
HttpCompressionStatic
Dynamic Content Compression
HttpCompressionDynamic
Management Tools 
IIS Management Console
ManagementConsole
IIS Management Scripts and Tools
ManagementScriptingTools
Management Service
AdminService
IIS 6 Management Compatibility
IIS Metabase Compatibility
Metabase
IIS 6 WMI Compatibility
WMICompatibility
IIS 6 Scripting Tools
LegacyScripts
IIS 6 Management Console
LegacySnapin
FTP Publishing Service
FTP Server
FTPServer
FTP Management snap-in
LegacySnapin
Windows Process Activation Service 
Process Model
ProcessModel
.NET Environment
NetFxEnvironment
Configuration APIs
WASConfigurationAPI

*/
void ModulesNT( CCFXRequest* pRequest )
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
	int iServer = pColumns->AddString( "Module" );
	int iState = pColumns->AddString( "State" );

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

	char *modules[]={
		"Web Server","W3SVC",
		"Common HTTP Features", "Static Content",
		/*
		StaticContent
		Default Document
		DefaultDocument
		Directory Browsing
		DirectoryBrowse
		HTTP Errors
		HttpErrors
		HTTP Redirection
		HttpRedirect
		Application Development Features 
		ASP.NET
		ASPNET
		.NET Extensibility
		NetFxExtensibility
		ASP
		ASP
		CGI
		CGI
		ISAPI Extensions
		ISAPIExtensions
		ISAPI Filters
		ISAPIFilter
		Server-Side Includes
		ServerSideInclude
		Health and Diagnostics 
		HTTP Logging
		HttpLogging
		Logging Tools
		LoggingLibraries
		Request Monitor
		RequestMonitor
		Tracing
		HttpTracing
		Custom Logging
		CustomLogging
		ODBC Logging
		ODBCLogging
		Security
		Basic Authentication
		BasicAuthentication
		Windows Authentication
		WindowsAuthentication
		Digest Authentication
		DigestAuthentication
		Client Certificate Mapping Authentication
		ClientCertificateMappingAuthentication
		IIS Client Certificate Mapping Authentication
		IISClientCertificateMappingAuthentication
		URL Authorization
		Authorization
		Request Filtering
		RequestFiltering
		IP and Domain Restrictions
		IPSecurity
		Performance Features 
		Static Content Compression
		HttpCompressionStatic
		Dynamic Content Compression
		HttpCompressionDynamic
		Management Tools 
		IIS Management Console
		ManagementConsole
		IIS Management Scripts and Tools
		ManagementScriptingTools
		Management Service
		AdminService
		IIS 6 Management Compatibility
		IIS Metabase Compatibility
		Metabase
		IIS 6 WMI Compatibility
		WMICompatibility
		IIS 6 Scripting Tools
		LegacyScripts
		IIS 6 Management Console
		LegacySnapin
		FTP Publishing Service
		FTP Server
		FTPServer
		FTP Management snap-in
		LegacySnapin
		Windows Process Activation Service 
		Process Model
		ProcessModel
		.NET Environment
		NetFxEnvironment
		Configuration APIs
		WASConfigurationAPI
		*/
	};

	/*
	HKEY hRegKey;
	char subkey[MAX_PATH+1];
	long ret;


	_snprintf(subkey,MAX_PATH,"Software\\Microsoft\\InetStp\\Components");

	bool bstate=false;
	ret=RegOpenKeyEx(HKEY_LOCAL_MACHINE,subkey,0,KEY_READ|KEY_WOW64_64KEY,&hRegKey);
	if(ret==ERROR_SUCCESS)
	{
		RegCloseKey(hRegKey);
		bstate=true;
	}

	//
	IADs *pADs=(IADs*)var.pdispVal;

	//
	CString strModule;
	CString strState; //ServerState

	//
	BSTR bstrName;
	hr2 = pADs->get_Name(&bstrName);
	log_COMError(__LINE__,hr2);
	if(SUCCEEDED(hr2)) strModule=bstrName;
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
	CString strState;
	if(bstate==true)
		strState="True";
	else
		strState="False";

	//
	int iRow = pQuery->AddRow();
	pQuery->SetData( iRow, iServer, strServer );

#ifdef _DEBUG
	pRequest->Write("</ul>DEBUGGING: end enum...<br>");
#endif					
	//
	hr2=pADs->Release();
	log_COMError(__LINE__,hr2);

	//
	SysFreeString(bstrState);
	SysFreeString(bstrModule);

#ifdef _DEBUG
	pRequest->Write("DEBUGGING: EXIT: ModulesNT...<br>");
#endif
	*/
}

// old trick to get the total number of lines of code
int getlines_modules() { return __LINE__+1; }
