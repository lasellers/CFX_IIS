// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////

//prototype
int getlines();

//
#define MAX_PREFMAXLEN			16384
#define MAX_UNC					12
#define MAX_NAME				256
#define MAX_TEXT				8192
#define MAX_DWORDEXPANSION		11
#define MAX_LDWORDEXPANSION		32
#define MAX_BOOLEANEXPANSION	4
#define MAX_TIMEDATE			128
#define MAX_PHONENUMBER			60
#define MAX_CN					64+3+1
#define MAX_SAM					20+1

#define MAX_ERRTEXT			 8192

// from iscnfg.h (platform dsk)
//1 (starting), 2 (started), 3 (stopping), 4 (stopped), 5 (pausing), 6 (paused), or 7 (continuing). 
//  Valid values for MD_SERVER_STATE
#define MD_SERVER_STATE_STARTING        1
#define MD_SERVER_STATE_STARTED         2
#define MD_SERVER_STATE_STOPPING        3
#define MD_SERVER_STATE_STOPPED         4
#define MD_SERVER_STATE_PAUSING         5
#define MD_SERVER_STATE_PAUSED          6
#define MD_SERVER_STATE_CONTINUING      7
