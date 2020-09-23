Attribute VB_Name = "MRasGlobals"
Option Explicit
'**********************************
'*     Constant Declarations      *
'**********************************
'constants needed for UDTs
Public Const UNLEN = 256
Public Const DNLEN = 15
Public Const PWLEN = 256

Public Const RAS_MaxDeviceType = 16
Public Const RAS_MaxPhoneNumber = 128
Public Const RAS_MaxIpAddress = 15
Public Const RAS_MaxIpxAddress = 21


Public Const RAS_MaxEntryName = 20
Public Const RAS_MaxDeviceName = 32
Public Const RAS_MaxCallbackNumber = 48

Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceName = 128
Public Const RAS95_MaxCallbackNumber = RAS_MaxPhoneNumber


Public Const RASP_Amb = &H10000
Public Const RASP_PppNbf = &H803F&
Public Const RASP_PppIpx = &H802B&
Public Const RASP_PppIp = &H8021&

'Other Constants
Public Const NETBIOS_NAME_LEN = 16

Public Const APINULL = 0&

Public Const VER_PLATFORM_WIN32s = 0&
Public Const VER_PLATFORM_WIN32_WINDOWS = 1&
Public Const VER_PLATFORM_WIN32_NT = 2&

'RASCONNSTATE enum
Public Const RASCS_PAUSED = &H1000&
Public Const RASCS_DONE = &H2000&

'begin enum
Public Const RASCS_OpenPort = 0&
Public Const RASCS_PortOpened = 1&
Public Const RASCS_ConnectDevice = 2&
Public Const RASCS_DeviceConnected = 3&
Public Const RASCS_AllDevicesConnected = 4&
Public Const RASCS_Authenticate = 5&
Public Const RASCS_AuthNotify = 6&
Public Const RASCS_AuthRetry = 7&
Public Const RASCS_AuthCallback = 8&
Public Const RASCS_AuthChangePassword = 9&
Public Const RASCS_AuthProject = 10&
Public Const RASCS_AuthLinkSpeed = 11&
Public Const RASCS_AuthAck = 12&
Public Const RASCS_ReAuthenticate = 13&
Public Const RASCS_Authenticated = 14&
Public Const RASCS_PrepareForCallback = 15&
Public Const RASCS_WaitForModemReset = 16&
Public Const RASCS_WaitForCallback = 17&
Public Const RASCS_Projected = 18&
 
Public Const RASCS_StartAuthentication = 19&    'Windows 95 only
Public Const RASCS_CallbackComplete = 20&        'Windows 95 only
Public Const RASCS_LogonNetwork = 21&            'Windows 95 only
 
Public Const RASCS_Interactive = RASCS_PAUSED
Public Const RASCS_RetryAuthentication = RASCS_PAUSED + 1&
Public Const RASCS_CallbackSetByCaller = RASCS_PAUSED + 2&
Public Const RASCS_PasswordExpired = RASCS_PAUSED + 3&
 
Public Const RASCS_Connected = RASCS_DONE
Public Const RASCS_Disconnected = RASCS_DONE + 1&
'end enum


'**********************************
'* User Defined Type Declarations *
'**********************************
'As a note VB subscripts are already +1 over C
Public Type RASDIALEXTENSIONS
    'set dwsize to 16
    dwSize As Long
    dwfOptions As Long
    hwndParent As Long
    reserved As Long
End Type

Public Type RASDIALPARAMS
    'set dwsize to 736 unless winver >= 400 then set to 1052
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
    szPhoneNumber(RAS_MaxPhoneNumber) As Byte
    szCallbackNumber(RAS_MaxCallbackNumber) As Byte
    szUserName(UNLEN) As Byte
    szPassword(PWLEN) As Byte
    szDomain(DNLEN) As Byte
End Type

Public Type RASDIALPARAMS95
    'set dwsize to 1052
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szPhoneNumber(RAS_MaxPhoneNumber) As Byte
    szCallbackNumber(RAS95_MaxCallbackNumber) As Byte
    szUserName(UNLEN) As Byte
    szPassword(PWLEN) As Byte
    szDomain(DNLEN) As Byte
End Type

Public Type RASCONN
    'set dwsize to 32
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MaxEntryName) As Byte
End Type

Public Type RASCONN95
    'set dwsize to 412
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Public Type RASENTRYNAME
    'set dwsize to 28 unless winver >= 400 then set to 264
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
End Type

Public Type RASENTRYNAME95
    'set dwsize to 264
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type

Public Type RASCONNSTATUS
    'set dwsize to 64 unless winver >= 400 then set to 288
    dwSize As Long
    rasconnstate As Long                            'RASCONNSTATE Enumeration
    dwError As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Public Type RASCONNSTATUS95
    'set dwsize to 160
    dwSize As Long
    rasconnstate As Long                            'RASCONNSTATE Enumeration
    dwError As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Public Type RASAMB
    'set dwsize to 28
    dwSize As Long
    dwError As Long
    szNetBiosError(NETBIOS_NAME_LEN) As Byte
    bLana As Byte
End Type

Public Type RASPPPNBF
    'set dwsize to 48
    dwSize As Long
    dwError As Long
    dwNetBiosError As Long
    szNetBiosError(NETBIOS_NAME_LEN) As Byte
    szWorkstationName(NETBIOS_NAME_LEN) As Byte
    bLana As Byte
End Type

Public Type RASPPPIPX
    'set dwsize to 32
    dwSize As Long
    dwError As Long
    szIpxAddress(RAS_MaxIpxAddress) As Byte
End Type

Public Type RASPPPIP
    'set dwsize to 40
    dwSize As Long
    dwError As Long
    szIpAddress(RAS_MaxIpAddress) As Byte
    szServerAddress(RAS_MaxIpAddress) As Byte
End Type


'**********************************
'*    WIN32 Type Declarations     *
'**********************************
'We have to determine the OS version
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type


'**********************************
'*   RAS Function Declarations    *
'**********************************
'I keep type checking wherever possible.
'Some functions need a ByVal sometimes and ByRef others. I declare ByRef and issue ByVal in function call.
'These functions should be good for all cases if the ByVal is added to the call wherever needed.
Public Declare Function RasEnumEntries Lib "RasApi32.DLL" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Public Declare Function RasCreatePhonebookEntry Lib "RasApi32.DLL" Alias "RasCreatePhonebookEntryA" (ByVal hWND As Long, ByVal lpszPhonebook As String) As Long
Public Declare Function RasEditPhonebookEntry Lib "RasApi32.DLL" Alias "RasEditPhonebookEntryA" (ByVal hWND As Long, ByVal lpszPhonebook As String, ByVal lpszEntryName As String) As Long


Public Declare Function RasDeleteEntry Lib "RasApi32.DLL" Alias "RasDeleteEntryA" (ByVal lpszPhonebook As String, ByVal lpszEntryName As String) As Long
Public Declare Function RasRenameEntry Lib "RasApi32.DLL" Alias "RasRenameEntryA" (ByVal lpszPhonebook As String, ByVal lpszOldEntryName As String, ByVal lpszNewEntryName As String) As Long
Public Declare Function RasValidateEntryName Lib "RasApi32.DLL" Alias "RasValidateEntryNameA" (ByVal lpszPhonebook As String, ByVal lpszEntryName As String) As Long


Public Declare Function RasGetEntryDialParams Lib "RasApi32.DLL" Alias "RasGetEntryDialParamsA" (ByVal lpszPhonebook As String, lprasdialparams As Any, lpfPassword As Long) As Long
Public Declare Function RasSetEntryDialParams Lib "RasApi32.DLL" Alias "RasSetEntryDialParamsA" (ByVal lpszPhonebook As String, lprasdialparams As Any, ByVal fRemovePassword As Long) As Long


Public Declare Function RasDial Lib "RasApi32.DLL" Alias "RasDialA" (lpRasDialExtensions As Any, ByVal lpszPhonebook As String, lprasdialparams As Any, ByVal dwNotifierType As Long, lpvNotifier As Long, lphRasConn As Long) As Long
Public Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasGetConnectStatus Lib "RasApi32.DLL" Alias "RasGetConnectStatusA" (ByVal hRasConn As Long, lpRASCONNSTATUS As Any) As Long
Public Declare Function RasGetErrorString Lib "RasApi32.DLL" Alias "RasGetErrorStringA" (ByVal uErrorValue As Long, ByVal lpszErrorString As String, ByVal cBufSize As Long) As Long
Public Declare Function RasGetProjectionInfo Lib "RasApi32.DLL" Alias "RasGetProjectionInfoA" (ByVal hRasConn As Long, ByVal rasprojection As Long, lpprojection As Any, lpcb As Long) As Long
Public Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

'**********************************
'*  WIN32 Function Declarations   *
'**********************************
'I use these all over the place so why duplicate declares
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'had to modify to fit my needs (Usually copying string to byte array because StrConv fails when array is not dynamic)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, ByVal lpString2 As String) As Long
Public Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


'**********************************
'*     RAS Error Return Codes     *
'**********************************
Public Const NOT_SUPPORTED = 120&

Public Const RASBASE = 600&
Public Const SUCCESS = 0&

Public Const PENDING = (RASBASE + 0)
Public Const ERROR_INVALID_PORT_HANDLE = (RASBASE + 1)
Public Const ERROR_PORT_ALREADY_OPEN = (RASBASE + 2)
Public Const ERROR_BUFFER_TOO_SMALL = (RASBASE + 3)
Public Const ERROR_WRONG_INFO_SPECIFIED = (RASBASE + 4)
Public Const ERROR_CANNOT_SET_PORT_INFO = (RASBASE + 5)
Public Const ERROR_PORT_NOT_CONNECTED = (RASBASE + 6)
Public Const ERROR_EVENT_INVALID = (RASBASE + 7)
Public Const ERROR_DEVICE_DOES_NOT_EXIST = (RASBASE + 8)
Public Const ERROR_DEVICETYPE_DOES_NOT_EXIST = (RASBASE + 9)
Public Const ERROR_BUFFER_INVALID = (RASBASE + 10)
Public Const ERROR_ROUTE_NOT_AVAILABLE = (RASBASE + 11)
Public Const ERROR_ROUTE_NOT_ALLOCATED = (RASBASE + 12)
Public Const ERROR_INVALID_COMPRESSION_SPECIFIED = (RASBASE + 13)
Public Const ERROR_OUT_OF_BUFFERS = (RASBASE + 14)
Public Const ERROR_PORT_NOT_FOUND = (RASBASE + 15)
Public Const ERROR_ASYNC_REQUEST_PENDING = (RASBASE + 16)
Public Const ERROR_ALREADY_DISCONNECTING = (RASBASE + 17)
Public Const ERROR_PORT_NOT_OPEN = (RASBASE + 18)
Public Const ERROR_PORT_DISCONNECTED = (RASBASE + 19)
Public Const ERROR_NO_ENDPOINTS = (RASBASE + 20)
Public Const ERROR_CANNOT_OPEN_PHONEBOOK = (RASBASE + 21)
Public Const ERROR_CANNOT_LOAD_PHONEBOOK = (RASBASE + 22)
Public Const ERROR_CANNOT_FIND_PHONEBOOK_ENTRY = (RASBASE + 23)
Public Const ERROR_CANNOT_WRITE_PHONEBOOK = (RASBASE + 24)
Public Const ERROR_CORRUPT_PHONEBOOK = (RASBASE + 25)
Public Const ERROR_CANNOT_LOAD_STRING = (RASBASE + 26)
Public Const ERROR_KEY_NOT_FOUND = (RASBASE + 27)
Public Const ERROR_DISCONNECTION = (RASBASE + 28)
Public Const ERROR_REMOTE_DISCONNECTION = (RASBASE + 29)
Public Const ERROR_HARDWARE_FAILURE = (RASBASE + 30)
Public Const ERROR_USER_DISCONNECTION = (RASBASE + 31)
Public Const ERROR_INVALID_SIZE = (RASBASE + 32)
Public Const ERROR_PORT_NOT_AVAILABLE = (RASBASE + 33)
Public Const ERROR_CANNOT_PROJECT_CLIENT = (RASBASE + 34)
Public Const ERROR_UNKNOWN = (RASBASE + 35)
Public Const ERROR_WRONG_DEVICE_ATTACHED = (RASBASE + 36)
Public Const ERROR_BAD_STRING = (RASBASE + 37)
Public Const ERROR_REQUEST_TIMEOUT = (RASBASE + 38)
Public Const ERROR_CANNOT_GET_LANA = (RASBASE + 39)
Public Const ERROR_NETBIOS_ERROR = (RASBASE + 40)
Public Const ERROR_SERVER_OUT_OF_RESOURCES = (RASBASE + 41)
Public Const ERROR_NAME_EXISTS_ON_NET = (RASBASE + 42)
Public Const ERROR_SERVER_GENERAL_NET_FAILURE = (RASBASE + 43)
Public Const WARNING_MSG_ALIAS_NOT_ADDED = (RASBASE + 44)
Public Const ERROR_AUTH_INTERNAL = (RASBASE + 45)
Public Const ERROR_RESTRICTED_LOGON_HOURS = (RASBASE + 46)
Public Const ERROR_ACCT_DISABLED = (RASBASE + 47)
Public Const ERROR_PASSWD_EXPIRED = (RASBASE + 48)
Public Const ERROR_NO_DIALIN_PERMISSION = (RASBASE + 49)
Public Const ERROR_SERVER_NOT_RESPONDING = (RASBASE + 50)
Public Const ERROR_FROM_DEVICE = (RASBASE + 51)
Public Const ERROR_UNRECOGNIZED_RESPONSE = (RASBASE + 52)
Public Const ERROR_MACRO_NOT_FOUND = (RASBASE + 53)
Public Const ERROR_MACRO_NOT_DEFINED = (RASBASE + 54)
Public Const ERROR_MESSAGE_MACRO_NOT_FOUND = (RASBASE + 55)
Public Const ERROR_DEFAULTOFF_MACRO_NOT_FOUND = (RASBASE + 56)
Public Const ERROR_FILE_COULD_NOT_BE_OPENED = (RASBASE + 57)
Public Const ERROR_DEVICENAME_TOO_LONG = (RASBASE + 58)
Public Const ERROR_DEVICENAME_NOT_FOUND = (RASBASE + 59)
Public Const ERROR_NO_RESPONSES = (RASBASE + 60)
Public Const ERROR_NO_COMMAND_FOUND = (RASBASE + 61)
Public Const ERROR_WRONG_KEY_SPECIFIED = (RASBASE + 62)
Public Const ERROR_UNKNOWN_DEVICE_TYPE = (RASBASE + 63)
Public Const ERROR_ALLOCATING_MEMORY = (RASBASE + 64)
Public Const ERROR_PORT_NOT_CONFIGURED = (RASBASE + 65)
Public Const ERROR_DEVICE_NOT_READY = (RASBASE + 66)
Public Const ERROR_READING_INI_FILE = (RASBASE + 67)
Public Const ERROR_NO_CONNECTION = (RASBASE + 68)
Public Const ERROR_BAD_USAGE_IN_INI_FILE = (RASBASE + 69)
Public Const ERROR_READING_SECTIONNAME = (RASBASE + 70)
Public Const ERROR_READING_DEVICETYPE = (RASBASE + 71)
Public Const ERROR_READING_DEVICENAME = (RASBASE + 72)
Public Const ERROR_READING_USAGE = (RASBASE + 73)
Public Const ERROR_READING_MAXCONNECTBPS = (RASBASE + 74)
Public Const ERROR_READING_MAXCARRIERBPS = (RASBASE + 75)
Public Const ERROR_LINE_BUSY = (RASBASE + 76)
Public Const ERROR_VOICE_ANSWER = (RASBASE + 77)
Public Const ERROR_NO_ANSWER = (RASBASE + 78)
Public Const ERROR_NO_CARRIER = (RASBASE + 79)
Public Const ERROR_NO_DIALTONE = (RASBASE + 80)
Public Const ERROR_IN_COMMAND = (RASBASE + 81)
Public Const ERROR_WRITING_SECTIONNAME = (RASBASE + 82)
Public Const ERROR_WRITING_DEVICETYPE = (RASBASE + 83)
Public Const ERROR_WRITING_DEVICENAME = (RASBASE + 84)
Public Const ERROR_WRITING_MAXCONNECTBPS = (RASBASE + 85)
Public Const ERROR_WRITING_MAXCARRIERBPS = (RASBASE + 86)
Public Const ERROR_WRITING_USAGE = (RASBASE + 87)
Public Const ERROR_WRITING_DEFAULTOFF = (RASBASE + 88)
Public Const ERROR_READING_DEFAULTOFF = (RASBASE + 89)
Public Const ERROR_EMPTY_INI_FILE = (RASBASE + 90)
Public Const ERROR_AUTHENTICATION_FAILURE = (RASBASE + 91)
Public Const ERROR_PORT_OR_DEVICE = (RASBASE + 92)
Public Const ERROR_NOT_BINARY_MACRO = (RASBASE + 93)
Public Const ERROR_DCB_NOT_FOUND = (RASBASE + 94)
Public Const ERROR_STATE_MACHINES_NOT_STARTED = (RASBASE + 95)
Public Const ERROR_STATE_MACHINES_ALREADY_STARTED = (RASBASE + 96)
Public Const ERROR_PARTIAL_RESPONSE_LOOPING = (RASBASE + 97)
Public Const ERROR_UNKNOWN_RESPONSE_KEY = (RASBASE + 98)
Public Const ERROR_RECV_BUF_FULL = (RASBASE + 99)
Public Const ERROR_CMD_TOO_LONG = (RASBASE + 100)
Public Const ERROR_UNSUPPORTED_BPS = (RASBASE + 101)
Public Const ERROR_UNEXPECTED_RESPONSE = (RASBASE + 102)
Public Const ERROR_INTERACTIVE_MODE = (RASBASE + 103)
Public Const ERROR_BAD_CALLBACK_NUMBER = (RASBASE + 104)
Public Const ERROR_INVALID_AUTH_STATE = (RASBASE + 105)
Public Const ERROR_WRITING_INITBPS = (RASBASE + 106)
Public Const ERROR_X25_DIAGNOSTIC = (RASBASE + 107)
Public Const ERROR_ACCT_EXPIRED = (RASBASE + 108)
Public Const ERROR_CHANGING_PASSWORD = (RASBASE + 109)
Public Const ERROR_OVERRUN = (RASBASE + 110)
Public Const ERROR_RASMAN_CANNOT_INITIALIZE = (RASBASE + 111)
Public Const ERROR_BIPLEX_PORT_NOT_AVAILABLE = (RASBASE + 112)
Public Const ERROR_NO_ACTIVE_ISDN_LINES = (RASBASE + 113)
Public Const ERROR_NO_ISDN_CHANNELS_AVAILABLE = (RASBASE + 114)
Public Const ERROR_TOO_MANY_LINE_ERRORS = (RASBASE + 115)
Public Const ERROR_IP_CONFIGURATION = (RASBASE + 116)
Public Const ERROR_NO_IP_ADDRESSES = (RASBASE + 117)
Public Const ERROR_PPP_TIMEOUT = (RASBASE + 118)
Public Const ERROR_PPP_REMOTE_TERMINATED = (RASBASE + 119)
Public Const ERROR_PPP_NO_PROTOCOLS_CONFIGURED = (RASBASE + 120)
Public Const ERROR_PPP_NO_RESPONSE = (RASBASE + 121)
Public Const ERROR_PPP_INVALID_PACKET = (RASBASE + 122)
Public Const ERROR_PHONE_NUMBER_TOO_LONG = (RASBASE + 123)
Public Const ERROR_IPXCP_NO_DIALOUT_CONFIGURED = (RASBASE + 124)
Public Const ERROR_IPXCP_NO_DIALIN_CONFIGURED = (RASBASE + 125)
Public Const ERROR_IPXCP_DIALOUT_ALREADY_ACTIVE = (RASBASE + 126)
Public Const ERROR_ACCESSING_TCPCFGDLL = (RASBASE + 127)
Public Const ERROR_NO_IP_RAS_ADAPTER = (RASBASE + 128)
Public Const ERROR_SLIP_REQUIRES_IP = (RASBASE + 129)
Public Const ERROR_PROJECTION_NOT_COMPLETE = (RASBASE + 130)
Public Const ERROR_PROTOCOL_NOT_CONFIGURED = (RASBASE + 131)
Public Const ERROR_PPP_NOT_CONVERGING = (RASBASE + 132)
Public Const ERROR_PPP_CP_REJECTED = (RASBASE + 133)
Public Const ERROR_PPP_LCP_TERMINATED = (RASBASE + 134)
Public Const ERROR_PPP_REQUIRED_ADDRESS_REJECTED = (RASBASE + 135)
Public Const ERROR_PPP_NCP_TERMINATED = (RASBASE + 136)
Public Const ERROR_PPP_LOOPBACK_DETECTED = (RASBASE + 137)
Public Const ERROR_PPP_NO_ADDRESS_ASSIGNED = (RASBASE + 138)
Public Const ERROR_CANNOT_USE_LOGON_CREDENTIALS = (RASBASE + 139)
Public Const ERROR_TAPI_CONFIGURATION = (RASBASE + 140)
Public Const ERROR_NO_LOCAL_ENCRYPTION = (RASBASE + 141)
Public Const ERROR_NO_REMOTE_ENCRYPTION = (RASBASE + 142)
Public Const ERROR_REMOTE_REQUIRES_ENCRYPTION = (RASBASE + 143)
Public Const ERROR_IPXCP_NET_NUMBER_CONFLICT = (RASBASE + 144)
Public Const ERROR_INVALID_SMM = (RASBASE + 145)
Public Const ERROR_SMM_UNINITIALIZED = (RASBASE + 146)
Public Const ERROR_NO_MAC_FOR_PORT = (RASBASE + 147)
Public Const ERROR_SMM_TIMEOUT = (RASBASE + 148)
Public Const ERROR_BAD_PHONE_NUMBER = (RASBASE + 149)
Public Const ERROR_WRONG_MODULE = (RASBASE + 150)
Public Const RASBASEEND = (RASBASE + 150)


'**********************************
'*  Public Variable Declarations  *
'**********************************
'I hate to do it, but this is the OS version variable. No client can touch it here and it saves a lot of code.
Public lngWindowVersion As Long
'I have to protect it, yet see it everywhere. Array for PhoneEntry Objects

Public arrPEntry() As PhoneEntry
Public arrConnection() As Connection
'Same goes for these
'set up variables for other objects
'They are not that big so we will keep them initialized
Public lpConnections As Connections
Public lpRASError As RASError
Public lpPhoneEntries As PhoneEntries
'Error object
'ErrorNumber property
Public lngRASErrorNumber As Long
'Description property
Public strRASDescription As String

'Flag so that I can stop client from updating
Public boolAllowUpdate As Boolean

Public Function fcnTrimNulls(ByVal strFullofNulls As String) As String

   'This function just gets rid of the Nulls that StrConv leaves on 95
   'passing like this is odd, but it works in the fewest lines
   'I had to add this to handle 95 after the fact
   If (InStr(strFullofNulls, Chr$(0))) Then fcnTrimNulls = Left$(strFullofNulls, InStr(strFullofNulls, Chr$(0)) - 1)
   
End Function

