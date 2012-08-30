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

// v 1.3 COM Error Reporting


void log_COMError(int line, HRESULT hr)
{
	if (FAILED(hr))
	{
		//
		int Sev=(hr>>30)&3;
		int C=(hr>>29)&1;
		int R=(hr>>28)&1;
		int Facility=(hr>>16)&4095;
		int Code=hr&65535;
		
		//
		CString buf;
		
		//
		CString strSev;
		switch(Sev)
		{
		case 0: strSev="SUCCESS"; break;
		case 1: strSev="INFORMATIONAL"; break;
		case 2: strSev="WARNING"; break;
		case 3: strSev="ERROR"; break;
		};
		CString strFacility;
		switch(Facility)
		{
		case FACILITY_WINDOWS: strFacility="WINDOWS"; break; //8
		case FACILITY_URT: strFacility="URT"; break; //19
		case FACILITY_STORAGE: strFacility="STORAGE"; break; //3
		//case FACILITY_SSPI: strFacility="SSPI"; break; //9
		case FACILITY_SCARD: strFacility="SCARD"; break; //16
		case FACILITY_SETUPAPI: strFacility="SETUPAPI"; break; //15
		case FACILITY_SECURITY: strFacility="SECURITY"; break; //9
		case FACILITY_RPC: strFacility="RPC"; break; //1
		case FACILITY_WIN32: strFacility="WIN32"; break; //7
		case FACILITY_CONTROL: strFacility="CONTROL"; break; // 10
		case FACILITY_NULL: strFacility="NULL"; break; //0
		case FACILITY_MSMQ: strFacility="MSMQ"; break; //14
		case FACILITY_MEDIASERVER: strFacility="MEDIASERVER"; break; //13
		case FACILITY_INTERNET: strFacility="INTERNET"; break; //12
		case FACILITY_ITF: strFacility="ITF"; break; //4
		case FACILITY_DISPATCH: strFacility="DISPATCH"; break; //2
		case FACILITY_COMPLUS: strFacility="COMPLUS"; break; //17
		case FACILITY_CERT: strFacility="CERT"; break; //11
		case FACILITY_ACS: strFacility="ACS"; break; //20
		case FACILITY_AAF: strFacility="AAF"; break; //18
		default: strFacility="UNKNOWN";
		};
		
		//
		buf.Format("ADSI Error 0x%x(%d) (Severity=%u %s Customer=%u Reserve=%u Facility=%u %s Code=0x%X(%d)) ",
			hr,
			hr,
			Sev,
			strSev,
			C,
			R,
			Facility,
			strFacility,
			Code,
			Code
			);
		log_Error(__LINE__,buf);
		
		//if facility is Win32, get the Win32 error 
		if (HRESULT_FACILITY(hr)==FACILITY_WIN32)
		{
			DWORD dwLastError;
			WCHAR szErrorBuf[MAX_ERRTEXT];
			WCHAR szNameBuf[MAX_ERRTEXT];
			//Get extended error value.
			HRESULT hr_return = S_OK;
			hr_return = ADsGetLastError( &dwLastError, szErrorBuf, MAX_ERRTEXT-2, szNameBuf, MAX_ERRTEXT-2);
			if (SUCCEEDED(hr_return))
			{
				buf.Format("LastError=0x%x(%d) Text=%ws Provider=%ws ", dwLastError,dwLastError, szErrorBuf, szNameBuf);
			}
			if(dwLastError==0)
			{
				LPVOID lpMsgBuf;
				FormatMessage(
					FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
					NULL,
					Code,
					MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
					(LPTSTR)&lpMsgBuf, 
					MAX_ERRTEXT-2,
					NULL
					);
				buf.Format(" %s ", lpMsgBuf);
				LocalFree( lpMsgBuf );
			}
			log_Error(__LINE__,buf);
		}
		
		log_appendline(line);
	}
}


//MK_E_SYNTAX 800401e4
//
//
//	Values are 32 bit values layed out as follows:
//
//	 3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
//	 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
//	+---+-+-+-----------------------+-------------------------------+
//	|Sev|C|R|	  Facility			|				Code			|
//	+---+-+-+-----------------------+-------------------------------+
//
//	where
//
//		Sev - is the severity code
//
//			00 - Success
//			01 - Informational
//			10 - Warning
//			11 - Error
//
//		C - is the Customer code flag
//
//		R - is a reserved bit
//
//		Facility - is the facility code
//
//		Code - is the facility's status code
