VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Protected Variables for Connections Collection
'Connection object Count for Count Property
Private intCCount As Integer
'********************************************************************

Public Property Get Count() As Integer

   'Seems trivial, but I may need to expand later
   Dim lngRetCode

   lngRetCode = fcnRASEnumConnections()
   If lngRetCode Then
      Err.Raise vbObjectError + lngRetCode, "Connections.Count Failure", "RAS Failure"
   Else
      Count = intCCount
   End If

End Property

Private Function fcnRASEnumConnections() As Long
 
   Dim lngRetCode As Long
   Dim lpcb As Long
   Dim lpcConnections As Long
   Dim intArraySize As Integer
   Dim intLooper As Integer
   
   'Putting a maximum of 256 Entries. If it fails then we resize
   intArraySize = 255
   If lngWindowVersion = 2 Then
      'We are running NT
      ReDim lprasconn(intArraySize) As RASCONN
      lprasconn(0).dwSize = 32
      lpcb = 256 * lprasconn(0).dwSize
      lngRetCode = RasEnumConnections(lprasconn(0), lpcb, lpcConnections)
   Else
      'We are running 95 (lngWindowVersion =1)
      ReDim lprasconn95(intArraySize) As RASCONN95
      lprasconn95(0).dwSize = 412
      lpcb = 256 * lprasconn95(0).dwSize
      lngRetCode = RasEnumConnections(lprasconn95(0), lpcb, lpcConnections)
   End If
   Select Case lngRetCode
      Case SUCCESS
         'resize array so that it is correct size based on return from function
         If lpcConnections > 0 Then
            ReDim arrConnection(lpcConnections - 1) As Connection
            If lngWindowVersion = 2 Then
                'running NT
                For intLooper = 0 To UBound(arrConnection())
                    Set arrConnection(intLooper) = New Connection
                    'allow entryname update
                    boolAllowUpdate = True
                    arrConnection(intLooper).hRasConn = lprasconn(intLooper).hRasConn
                    arrConnection(intLooper).EntryName = fcnTrimNulls(StrConv(lprasconn(intLooper).szEntryName, vbUnicode))
                    arrConnection(intLooper).Index = intLooper
                    boolAllowUpdate = False
                Next intLooper
            Else
                    'running 95
                    For intLooper = 0 To UBound(arrConnection())
                    Set arrConnection(intLooper) = New Connection
                    'allow entryname update
                    boolAllowUpdate = True
                    arrConnection(intLooper).hRasConn = lprasconn95(intLooper).hRasConn
                    arrConnection(intLooper).EntryName = fcnTrimNulls(StrConv(lprasconn95(intLooper).szEntryName, vbUnicode))
                    arrConnection(intLooper).Index = intLooper
                    boolAllowUpdate = False
                Next intLooper
            End If
         End If
         'set the Connections.Count
         'I doubt that this will ever fail, but I have been wrong
         intCCount = CInt(lpcConnections)
         fcnRASEnumConnections = 0
      Case ERROR_BUFFER_TOO_SMALL
         'Make buffers bigger and try again
         If lngWindowVersion = 2 Then
            'running NT
            intArraySize = lpcb / lprasconn(0).dwSize
         Else
            'running 95
            intArraySize = lpcb / lprasconn95(0).dwSize
         End If
      Case Else
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         fcnRASEnumConnections = lngRetCode
   End Select
  
End Function

Public Function AddConnection(strNewEntryName As String, strNewPhoneNumber As String, strNewCallbackNumber As String, strNewUsername As String, strNewPassword As String, strNewDomain As String, boolAsync As Boolean) As Connection

   'Kind of funky Add, but that is the way RAS works.
   'Since there is no way of guaranteeing the object return, I return the handle to the Connection
   Dim lngRetCode As Long
   Dim hRasConn As Long
   Dim lngRetlstrcpy As Long
   Dim intLooper As Integer
   Dim lngRetHangUp As Long
   
      'Doing RASDIAL call Async in a VB sort of way. Since we can not define a callback, we call
      'The API and pass in the handle of a modal form that we have in the server. We then poll the
      'connection with RASGetConnectionStatus until we see that we are connected
      If lngWindowVersion = 2 Then
         'We are running NT
         Dim lprasdialparams As RASDIALPARAMS
         lprasdialparams.dwSize = 736
         'Using lstrcpy because StrConv fails. I let VB convert the string and fill the array
         'I should theoretically be error checking here, but chances of failure are slim and
         'I should catch it because RASDIAL will fail
         lngRetlstrcpy = lstrcpy(lprasdialparams.szEntryName(0), strNewEntryName)
         lngRetlstrcpy = lstrcpy(lprasdialparams.szPhoneNumber(0), strNewPhoneNumber)
         lngRetlstrcpy = lstrcpy(lprasdialparams.szCallbackNumber(0), strNewCallbackNumber)
         lngRetlstrcpy = lstrcpy(lprasdialparams.szUserName(0), strNewUsername)
         lngRetlstrcpy = lstrcpy(lprasdialparams.szPassword(0), strNewPassword)
         lngRetlstrcpy = lstrcpy(lprasdialparams.szDomain(0), strNewDomain)
         'Call RASDial
         If boolAsync Then
            'Asyncronous and ignoring RASDIALEXTENSIONS.
            'So that HWND is valid
            'Load frmAsyncDial
            lngRetCode = RasDial(ByVal APINULL, vbNullString, lprasdialparams, &HFFFFFFFF, FDialup.hWND, hRasConn)
         Else
            'Syncronous and ignoring RASDIALEXTENSIONS.
            Screen.MousePointer = vbHourglass
            lngRetCode = RasDial(ByVal APINULL, vbNullString, lprasdialparams, APINULL, ByVal APINULL, hRasConn)
            Screen.MousePointer = vbDefault
         End If
         'Test for failure and raise error if so
         If lngRetCode Then
            lngRASErrorNumber = lngRetCode
            strRASDescription = lpRASError.fcnRASErrorString()
            lngRetHangUp = RasHangUp(hRasConn)
            Err.Raise vbObjectError + lngRetCode, "Connections AddConnection Failed", "RAS Failure"
         Else
            'Return the handle to the connection to the client if they need it. See below
            'Just giving time so that Connections is updated properly
            DoEvents
         End If
      Else
         'We are running 95 (lngWindowVersion =1)
         Dim lprasdialparams95 As RASDIALPARAMS95
         lprasdialparams95.dwSize = 1052
         'Using lstrcpy because StrConv fails. I let VB convert the string and fill the array
         'I should theoretically be error checking here, but chances of failure are slim and
         'I should catch it because RASDIAL will fail
         lngRetlstrcpy = lstrcpy(lprasdialparams95.szEntryName(0), strNewEntryName)
         lngRetlstrcpy = lstrcpy(lprasdialparams95.szPhoneNumber(0), strNewPhoneNumber)
         lngRetlstrcpy = lstrcpy(lprasdialparams95.szCallbackNumber(0), strNewCallbackNumber)
         lngRetlstrcpy = lstrcpy(lprasdialparams95.szUserName(0), strNewUsername)
         lngRetlstrcpy = lstrcpy(lprasdialparams95.szPassword(0), strNewPassword)
         lngRetlstrcpy = lstrcpy(lprasdialparams95.szDomain(0), strNewDomain)
         'Call RASDial
         If boolAsync Then
            'Asyncronous and ignoring RASDIALEXTENSIONS.
            'So that HWND is valid
            'Load frmAsyncDial
            'just to be sure
            DoEvents
            lngRetCode = RasDial(ByVal APINULL, vbNullString, lprasdialparams95, &HFFFFFFFF, FDialup.hWND, hRasConn)
         Else
            'Syncronous and ignoring RASDIALEXTENSIONS.
            Screen.MousePointer = vbHourglass
            lngRetCode = RasDial(ByVal APINULL, vbNullString, lprasdialparams95, APINULL, ByVal APINULL, hRasConn)
            Screen.MousePointer = vbDefault
         End If
         'Test for failure and raise error if so
         If lngRetCode Then
            lngRASErrorNumber = lngRetCode
            strRASDescription = lpRASError.fcnRASErrorString()
            lngRetHangUp = RasHangUp(hRasConn)
            Err.Raise vbObjectError + lngRetCode, "Connections AddConnection Failed", "RAS Failure"
         Else
            'Return the handle to the connection to the client if they need it. See below
            'Just giving time so that Connections is updated properly
            DoEvents
         End If
      End If
      
      'Return the handle to the connection to the client if they need it
      'Enumerate the connections so that we can return correct handle to connection object
      'This is the only way I can see to guarantee object without complex algorithms
      If boolAsync Then
         'this is sort of a kludge, but I am setting the Tag of the form to the hRasConn
         'so that I can reference it in the form without having to use public variables
         FDialup.Tag = Hex$(hRasConn)
         'show the async form so that processing stops here and get RAS_EVENTS in form
         'frmAsyncDial.Show 1
      Else
         'nothing to do if synchronous
      End If
     'Refresh and return the function to the client
      lngRetCode = fcnRASEnumConnections()
      If lngRetCode Then Err.Raise vbObjectError + lngRetCode, "Connections AddConnection Failed", "RAS Failure"
      For intLooper = 0 To intCCount - 1
         If hRasConn = arrConnection(intLooper).hRasConn Then
            Set AddConnection = arrConnection(intLooper)
         Else
            Set AddConnection = Nothing
         End If
      Next intLooper
      
End Function

Public Sub RemoveConnection(lngIndexToRem As Long)

   Dim lngRetCode As Long
   Dim hRasConnToRem As Long
        
   'get the hRasConn for the given index
   hRasConnToRem = arrConnection(lngIndexToRem).hRasConn
   
   'Call RASHangUp
   lngRetCode = RasHangUp(hRasConnToRem)
   If lngRetCode Then
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         Err.Raise vbObjectError + lngRetCode, "Connections RemoveConnection Failed", "RAS Failure"
      Else
         'reindex the connections
         lngRetCode = fcnRASEnumConnections()
         If lngRetCode Then Err.Raise vbObjectError + lngRetCode, "Connections RemoveConnection Failed", "RAS Failure"
   End If

End Sub

Private Sub Class_Initialize()

   Dim lngSuccess As Long
   
   'initialize Phone book collection
   'throw away error we will catch later
   lngSuccess = fcnRASEnumConnections()

End Sub

