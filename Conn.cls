VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'protected variables for connection class
Private intIndex As Integer
Private lnghRasConn As Long
Private strEntryName As String
Private lngRASConnState As Long
Private strDeviceType As String
Private strDeviceName As String
Private strPhoneNumber As String
Private strCallbackNumber As String
Private strUserName As String
Private strDomain As String
Private strPassword As String
Private strWorkstationName As String
Private strIPXAddress As String
Private strIPAddress As String
Private bytLANA As Byte

Public Property Get hRasConn() As Long

   hRasConn = lnghRasConn

End Property

Public Property Let hRasConn(hNewConn As Long)

   Dim lngRetCode As Long

   If boolAllowUpdate Then
      lnghRasConn = hNewConn
      'These are both extraneous so if they fail I just ignore them
      lngRetCode = fcnRASGetConnectionStatus()
      lngRetCode = fcnRasGetProjectionInfo()
   Else
      lngRASErrorNumber = 1111
      strRASDescription = "Property Not Updateable"
      Err.Raise vbObjectError + 1111, "Property Not Updateable", "RAS Failure"
   End If

End Property

Public Property Get EntryName() As String

   EntryName = strEntryName

End Property

Public Property Let EntryName(strNewName As String)

   Dim lngRetCode As Long

   'Test flag to be sure that it is not client
   If boolAllowUpdate Then
      strEntryName = strNewName
      If lngRetCode Then
         Err.Raise vbObjectError + lngRetCode, "EntryName Set Failed", "RAS Failure"
      End If
   Else
      lngRASErrorNumber = 1111
      strRASDescription = "Property Not Updateable"
      Err.Raise vbObjectError + 1111, "Property Not Updateable", "RAS Failure"
   End If


End Property

Public Property Get Index() As Integer

   Index = intIndex

End Property

Public Property Let Index(intNewIndex As Integer)

 If boolAllowUpdate Then
      intIndex = intNewIndex
   Else
      lngRASErrorNumber = 1111
      strRASDescription = "Property Not Updateable"
      Err.Raise vbObjectError + 1111, "Property Not Updateable", "RAS Failure"
   End If

End Property

Private Function fcnRASGetConnectionStatus() As Long

   Dim lngRetCode As Long
    
   If lngWindowVersion = 2 Then
      'using NT
      Dim lpRASCONNSTATUS As RASCONNSTATUS
      lpRASCONNSTATUS.dwSize = 64
      lngRetCode = RasGetConnectStatus(lnghRasConn, lpRASCONNSTATUS)
      If lngRetCode Then
         strDeviceName = "Not Available"
         strDeviceType = "Not Available"
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         fcnRASGetConnectionStatus = lngRetCode
      Else
         'success
         lngRASConnState = lpRASCONNSTATUS.rasconnstate
         strDeviceName = fcnTrimNulls(StrConv(lpRASCONNSTATUS.szDeviceName(), vbUnicode))
         strDeviceType = fcnTrimNulls(StrConv(lpRASCONNSTATUS.szDeviceType(), vbUnicode))
         fcnRASGetConnectionStatus = 0
      End If
   Else
      'using 95
      Dim lpRASCONNSTATUS95 As RASCONNSTATUS95
      lpRASCONNSTATUS95.dwSize = 160
      lngRetCode = RasGetConnectStatus(lnghRasConn, lpRASCONNSTATUS95)
      If lngRetCode Then
         strDeviceName = "Not Available"
         strDeviceType = "Not Available"
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         fcnRASGetConnectionStatus = lngRetCode
      Else
         'success
         lngRASConnState = lpRASCONNSTATUS95.rasconnstate
         strDeviceName = fcnTrimNulls(StrConv(lpRASCONNSTATUS95.szDeviceName(), vbUnicode))
         strDeviceType = fcnTrimNulls(StrConv(lpRASCONNSTATUS95.szDeviceType(), vbUnicode))
         fcnRASGetConnectionStatus = 0
      End If
   End If

End Function


Public Property Get State() As Long
   
   State = lngRASConnState
   
End Property


Public Property Get DeviceType() As String

   DeviceType = strDeviceType

End Property



Public Property Get DeviceName() As String

   DeviceName = strDeviceName

End Property

Public Property Get LANA() As Byte

   LANA = bytLANA

End Property

Public Property Get WorkstationName() As String

   WorkstationName = strWorkstationName

End Property

Public Property Get IPXAddress() As String

   IPXAddress = strIPXAddress

End Property

Public Property Get IPAddress() As String

   IPAddress = strIPAddress

End Property

Private Function fcnRasGetProjectionInfo() As Long

   Dim lngRetCode As Long
   Dim lprasamb As RASAMB
   Dim lpraspppnbf As RASPPPNBF
   Dim lpraspppipx As RASPPPIPX
   Dim lpraspppip As RASPPPIP
   Dim lpcb As Long
   Dim rasprojection As Long
   
   'this will only work for PPP connections
   'start with RASAMB
   rasprojection = RASP_Amb
   lprasamb.dwSize = 28
   lpcb = 28
   lngRetCode = RasGetProjectionInfo(lnghRasConn, rasprojection, lprasamb, lpcb)
   If lngRetCode Then
      bytLANA = 0
      lngRASErrorNumber = lngRetCode
      strRASDescription = lpRASError.fcnRASErrorString()
      fcnRasGetProjectionInfo = lngRetCode
   Else
      bytLANA = lprasamb.bLana
      fcnRasGetProjectionInfo = 0
   End If
   'Net BIOS
   rasprojection = RASP_PppNbf
   lpraspppnbf.dwSize = 48
   lpcb = 48
   lngRetCode = RasGetProjectionInfo(lnghRasConn, rasprojection, lpraspppnbf, lpcb)
   If lngRetCode Then
      bytLANA = 0
      strWorkstationName = "Not Available"
      lngRASErrorNumber = lngRetCode
      strRASDescription = lpRASError.fcnRASErrorString()
      fcnRasGetProjectionInfo = lngRetCode
   Else
      bytLANA = lpraspppnbf.bLana
      strWorkstationName = fcnTrimNulls(StrConv(lpraspppnbf.szWorkstationName, vbUnicode))
      fcnRasGetProjectionInfo = 0
   End If
   'IPX
   rasprojection = RASP_PppIpx
   lpraspppipx.dwSize = 32
   lpcb = 32
   lngRetCode = RasGetProjectionInfo(lnghRasConn, rasprojection, lpraspppipx, lpcb)
   If lngRetCode Then
      strIPXAddress = "Not Available"
      lngRASErrorNumber = lngRetCode
      strRASDescription = lpRASError.fcnRASErrorString()
      fcnRasGetProjectionInfo = lngRetCode
   Else
      strIPXAddress = fcnTrimNulls(StrConv(lpraspppipx.szIpxAddress, vbUnicode))
      fcnRasGetProjectionInfo = 0
   End If
   'TCP/IP
   rasprojection = RASP_PppIp
   lpraspppip.dwSize = 40
   lpcb = 40
   lngRetCode = RasGetProjectionInfo(lnghRasConn, rasprojection, lpraspppip, lpcb)
   If lngRetCode Then
      strIPAddress = "Not Available"
      lngRASErrorNumber = lngRetCode
      strRASDescription = lpRASError.fcnRASErrorString()
      fcnRasGetProjectionInfo = lngRetCode
   Else
      strIPAddress = fcnTrimNulls(StrConv(lpraspppip.szIpAddress, vbUnicode))
      fcnRasGetProjectionInfo = 0
   End If
   
End Function

