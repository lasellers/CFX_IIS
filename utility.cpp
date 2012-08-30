/* #ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif */

#include "stdafx.h"		// Standard MFC libraries
//#include "activeds.h"
#include "afxpriv.h"	// MFC Unicode conversion macros

#include "cfx.h" //..\..\classes\cfx.h"		// CFX Custom Tag API

//#include "basetsd.h"

//#include "lmaccess.h" // UNLEN LM20_PWLEN

#include "wchar.h"

// bstr
#include <objidl.h>
#include <comdef.h>

// adsi
//#include <iads.h>
//#include <adshlp.h>

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

//
// get 
/*
for bindings, securebindings, httperrors, HTTPCustomHeaders, etc.
Take the variant array recieved from get and makes it into a semicolon delimited string
*/
int VARIANTARRAYtoCString (VARIANT var, CString& str, CCFXRequest* pRequest, char seperator )
{
	USES_CONVERSION;
	
	//
	HRESULT hr;
	
	CString e;
	
	//
	LONG lstart, lend;
	SAFEARRAY *sa = V_ARRAY( &var );
	
	//
	hr = SafeArrayGetLBound( sa, 1, &lstart );
	log_COMError(__LINE__,hr);
	hr = SafeArrayGetUBound( sa, 1, &lend );
	log_COMError(__LINE__,hr);
	
	//
	VARIANT varItem;
	VariantInit(&varItem);
	
	//
	for ( long idx=lstart; idx <= lend; idx++ )
	{
		hr = SafeArrayGetElement( sa, &idx, &varItem );
		e+=V_BSTR(&varItem);
		VariantClear(&varItem);
		if(idx!=lend) e+=seperator;
	}
	
	//
	str=e;
	
	// return count
	return (lend-lstart+1);
}


//
// put
int CStringtoVARIANTARRAY(CString str, VARIANT *var, CCFXRequest* pRequest, char seperator)
{
	USES_CONVERSION;
	
	HRESULT hr=S_OK;
	
	//
	int buflen=str.GetLength()*2;
	char *s=new char[buflen];
	sprintf(s,"%s",str);
	
	//
	long count=0;
	{
		char *p=s;
		while(*p!=0)
		{
			if(*p==seperator)
			{
				*p=0;
				count++;
			}
			p++;
		};
		if(buflen>0) count++;
	}
	
	//
	if(count>0)
	{
		//
		SAFEARRAYBOUND saBound;
		saBound.lLbound=0;
		saBound.cElements=count;
		
		//
		SAFEARRAY *pSA = SafeArrayCreate( VT_VARIANT, 1, &saBound );
		if(pSA==NULL)
		{
			log_Error(__LINE__,"SAFEARRAY failure");
		}
		else
		{
			//
			char *m=s;
			for(long i=0;i<(long)saBound.cElements;i++)
			{
				//
				VARIANT varItem;
				VariantInit(&varItem);
				
				V_BSTR(&varItem) = SysAllocString(A2W(m));
				V_VT(&varItem) = VT_BSTR;
				
				SafeArrayPutElement(pSA, &i, &varItem);
				VariantClear(&varItem);
				
				//
				while(*m!=NULL)
					m++;
				m++;
			}
			
			//
			VariantClear(var);
			V_VT(var) = VT_VARIANT | VT_ARRAY;
			V_ARRAY(var) = pSA;
			
			//
			//			hr=SafeArrayDestroy(pSA);
			//			log_COMError(__LINE__,hr);
		}
	}
	
	//
	delete [] s;
	
	return count;
}
