//#include "cfx.h" //..\..\classes\cfx.h"		// CFX Custom Tag API

// v 1.3 COM Error Reporting

//
#define FACILITY_WINDOWS                 8
#define FACILITY_URT                     19
#define FACILITY_STORAGE                 3
#define FACILITY_SSPI                    9
#define FACILITY_SCARD                   16
#define FACILITY_SETUPAPI                15
#define FACILITY_SECURITY                9
#define FACILITY_RPC                     1
#define FACILITY_WIN32                   7
#define FACILITY_CONTROL                 10
#define FACILITY_NULL                    0
#define FACILITY_MSMQ                    14
#define FACILITY_MEDIASERVER             13
#define FACILITY_INTERNET                12
#define FACILITY_ITF                     4
#define FACILITY_DISPATCH                2
#define FACILITY_COMPLUS                 17
#define FACILITY_CERT                    11
#define FACILITY_ACS                     20
#define FACILITY_AAF                     18

void log_COMError(int line, HRESULT hr);
