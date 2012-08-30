/*
activeds.h
#define UNICODE
#define _UNICODE

ms's adsi sdk.zip

adsisdk tools/options/directories:
inc/
lib/ adsiid.lib and activeds.lib


NOTE TO SELF: for next version integrate redundancies into subfunctions
*/

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

#include "ftp.h"
#include "webserver.h"
#include "webdirectory.h"
#include "smtpserver.h"
#include "nntpserver.h"
#include "services.h"
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

//
void ProcessTagRequest( CCFXRequest* pRequest ) 
{
	try
	{
		::CoInitialize(NULL);

		// initialize the error stream (blank it)
		empty_log();

		// the working string to format data with
		CString tmp;

		//start the clock ticking to time the tag
		CTime starttick = CTime::GetCurrentTime();

		//
		CString strQuery = pRequest->GetAttribute( "QUERY" );
		if(!strQuery.IsEmpty())
			set_query_variable(strQuery);
		else
			set_query_variable("IIS");
		pRequest->SetVariable("IISQuery", get_query_variable());
		pRequest->SetVariable( get_query_variable(), "");

		//
		pRequest->SetVariable( "IIS", "" );
		pRequest->SetVariable( "IISError", "" );

		pRequest->SetVariable( "IISAction", "" );
		pRequest->SetVariable( "IISSeconds", "0" );

		//
		pRequest->SetVariable( "IISServer", "" );
		pRequest->SetVariable( "IISDirectory", "" );
		pRequest->SetVariable( "IISPath", "" );

		//
		/*		if(IHTKPassword(pRequest))
		{
		pRequest->SetVariable( "IISError", "Shared environment password protection failure. See IHTKPassword." );
		return;
		}
		*/

		//
		CString strAction = pRequest->GetAttribute( "ACTION" );
		pRequest->SetVariable("IISAction", strAction);

		// about
		if ( !strAction.CompareNoCase( "ABOUT" ) )
			AboutNT( pRequest );

		// site services
		else if ( !strAction.CompareNoCase( "STOP" ) )
			StopStartPauseContinueNT( pRequest, "Stop" );
		else if ( !strAction.CompareNoCase( "START" ) )
			StopStartPauseContinueNT( pRequest, "Start" );
		else if ( !strAction.CompareNoCase( "PAUSE" ) )
			StopStartPauseContinueNT( pRequest, "Pause" );
		else if ( !strAction.CompareNoCase( "CONTINUE" ) )
			StopStartPauseContinueNT( pRequest, "Continue" );

		// ftp server
		else if ( !strAction.CompareNoCase( "FTPSERVERS" ) )
			FTPServersNT( pRequest );
		else if ( !strAction.CompareNoCase( "FTPSERVER" ) )
			FTPServerNT( pRequest );
		else if ( !strAction.CompareNoCase( "ADDFTPSERVER" ) )
			AddFTPServerNT( pRequest );
		else if ( !strAction.CompareNoCase( "DELETEFTPSERVER" ) )
			DeleteFTPServerNT( pRequest );

		//ftp directory
		else if ( !strAction.CompareNoCase( "FTPDIRECTORIES" ) )
			FTPDirectoriesNT( pRequest );
		else if ( !strAction.CompareNoCase( "ADDFTPDIRECTORY" ) )
			AddFTPDirectoryNT( pRequest );
		else if ( !strAction.CompareNoCase( "DELETEFTPDIRECTORY" ) )
			DeleteFTPDirectoryNT( pRequest );

		// web server
		else if ( !strAction.CompareNoCase( "WEBSERVERS" ) )
			WebServersNT( pRequest );
		else if ( !strAction.CompareNoCase( "WEBSERVER" ) )
			WebServerNT( pRequest );
		else if ( !strAction.CompareNoCase( "ADDWEBSERVER" ) )
			AddWebServerNT( pRequest );
		else if ( !strAction.CompareNoCase( "DELETEWEBSERVER" ) )
			DeleteWebServerNT( pRequest );

		//web directory
		else if ( !strAction.CompareNoCase( "WEBDIRECTORIES" ) )
			WebDirectoriesNT( pRequest );
		else if ( !strAction.CompareNoCase( "ADDWEBDIRECTORY" ) )
			AddWebDirectoryNT( pRequest );
		else if ( !strAction.CompareNoCase( "DELETEWEBDIRECTORY" ) )
			DeleteWebDirectoryNT( pRequest );

		// smtp
		else if ( !strAction.CompareNoCase( "SMTPSERVERS" ) )
			SMTPServersNT( pRequest );
		else if ( !strAction.CompareNoCase( "ADDSMTPSERVER" ) )
			;//AddSMTPServerNT( pRequest );
		else if ( !strAction.CompareNoCase( "DELETESMTPSERVER" ) )
			;//DeleteSMTPServerNT( pRequest );

		// nntp
		else if ( !strAction.CompareNoCase( "NNTPSERVERS" ) )
			NNTPServersNT( pRequest );
		else if ( !strAction.CompareNoCase( "ADDNNTPSERVER" ) )
			;//AddNNTPServerNT( pRequest );
		else if ( !strAction.CompareNoCase( "DELETENNTPSERVER" ) )
			;//DeleteNNTPServerNT( pRequest );
		//
//		http://learn.iis.net/page.aspx/135/discover-installed-components/

		else if ( !strAction.CompareNoCase( "MODULES" ) )
			;//ModulesNT( pRequest );
		//
		else
		{
			if(!strAction.IsEmpty())
			{
				tmp.Format("An invalid ACTION attribute (%s) was specified. ",strAction);
				log_Error(__LINE__,tmp);
			}
		}

		//echo
		pRequest->SetVariable("IISError",get_log());

		//time
		CTime endtick = CTime::GetCurrentTime();
		CTimeSpan span=endtick-starttick;
		CString seconds;
		seconds.Format("%u",span.GetTotalSeconds());
		pRequest->SetVariable( "IISSeconds", seconds);

		::CoUninitialize();
	}

	// Catch Cold Fusion exceptions & re-raise them
	catch( CCFXException* e )
	{
		pRequest->ReThrowException( e );
	}

	// Catch ALL other exceptions and throw them as 
	// Cold Fusion exceptions (DO NOT REMOVE! -- 
	// this prevents the server from crashing in 
	// case of an unexpected exception)
	catch( ... )
	{
		pRequest->ThrowException(
			"Error occurred in tag CFX_IIS",
			"Unexpected and completely unwanted error occurred while processing tag."
			);
	}
}

// old trick to get the total number of lines of code
int getlines() { return __LINE__+1 + getlines_ftp() + getlines_nntpserver() + getlines_services() + getlines_smtpserver() + getlines_webdirectory() + getlines_webserver() + getlines_modules(); }
