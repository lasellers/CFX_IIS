// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////

//v1.1 August 30th 2004 fixed a crash under cfmx 6.1 + apache 2
//v1.0

// returns TRUE if password-protected and an early abort of tag is required.
int IHTKPassword(CCFXRequest* pRequest)
{
	bool ret=FALSE;
    long result;
    HKEY key;

    result = RegOpenKeyEx(
		HKEY_LOCAL_MACHINE,
		"SOFTWARE\\Intrafoundation\\serial",
		0,
		KEY_READ,
		&key
		);
    if (result == ERROR_SUCCESS)
	{
		DWORD type;
		DWORD bufsize;
		char buf[512];

		buf[0]='\0';
		bufsize = sizeof(buf);
		result = RegQueryValueEx(key,"", NULL, &type, (LPBYTE) buf, &bufsize);
		if (result == ERROR_SUCCESS)
		{
			LPCSTR lpcstrIHTKPassword = pRequest->GetAttribute( "IHTKPASSWORD" );
			if(lpcstrIHTKPassword!=NULL && strcmp(lpcstrIHTKPassword,buf)==0)
				;
			else
				ret=TRUE; //password protected
		}
		RegCloseKey(key);
	}

	return ret;
}
