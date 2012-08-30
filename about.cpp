// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////

#include "project.h"

// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
const char *umDescription="CFX_IIS. Written by Lewis A. Sellers, Copyright (c) 1999, 2000, 2001, 2003, 2004, 2012. Intrafoundation Software at http://www.intrafoundation.com.";
const char *umVersion="1.19a";
const char *umQuality="gamma";
//const char *umSubVersion="";

// v1.3

void AboutNT( CCFXRequest* pRequest )
{
	// fields
	CCFXStringSet* pColumns = pRequest->CreateStringSet();
	int iDescription = pColumns->AddString( "Description" );
	int iVersion = pColumns->AddString( "Version" );
	int iQuality = pColumns->AddString( "Quality" );
	int iSubVersion = pColumns->AddString( "SubVersion" );
	int iLines = pColumns->AddString( "Lines" );
//	int iSerialNumber = pColumns->AddString( "SerialNumber" );
	int iBuildDate = pColumns->AddString( "BuildDate" );
	
//	int iEvaluation= pColumns->AddString( "Evaluation" );
//	int iExpirationDate= pColumns->AddString( "ExpirationDate" );
	CCFXQuery* pQuery = pRequest->AddQuery( UMQuery, pColumns );
	
	//
	CString tmp;
	
	//
	int iRow = pQuery->AddRow();
	char sz[256];
	_snprintf(sz,255,"OPEN-SOURCED. %s",umDescription);
	pQuery->SetData( iRow, iDescription, sz );
	
	pQuery->SetData( iRow, iVersion, umVersion );
	pQuery->SetData( iRow, iQuality, umQuality );
//	pQuery->SetData( iRow, iSubVersion, umSubVersion );
	
	CString d;
	d.Format("%s",__DATE__); //Mmm dd yyyy
	int m=0;
	CString s=d.Mid(0,3);
	if(!s.CompareNoCase("Jan")) m=1;
	else if(!s.CompareNoCase("Feb")) m=2;
	else if(!s.CompareNoCase("Mar")) m=3;
	else if(!s.CompareNoCase("Apr")) m=4;
	else if(!s.CompareNoCase("May")) m=5;
	else if(!s.CompareNoCase("Jun")) m=6;
	else if(!s.CompareNoCase("Jul")) m=7;
	else if(!s.CompareNoCase("Aug")) m=8;
	else if(!s.CompareNoCase("Sep")) m=9;
	else if(!s.CompareNoCase("Oct")) m=10;
	else if(!s.CompareNoCase("Nov")) m=11;
	else if(!s.CompareNoCase("Dec")) m=12;
	tmp.Format("{ts '%04d-%02d-%02d %02d:%02d:%02d'}",
		atoi(d.Mid(7,4)), //year
		m, //month
		atoi(d.Mid(4,2)), //day
		0,
		0,
		0
		);
	pQuery->SetData( iRow, iBuildDate, tmp);
	
	tmp.Format("%d",getlines());
	pQuery->SetData( iRow, iLines, tmp);
	
//	pQuery->SetData( iRow, iSerialNumber, umSerialNumber );
	
//	tmp.Format("%d",_MSC_VER);
	
//	pQuery->SetData( iRow, iEvaluation, "0" );
//	pQuery->SetData( iRow, iExpirationDate, "" );
}
