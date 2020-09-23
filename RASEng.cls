VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RASEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Set up declares used in this object

'*************************************************************************************************************************

'set up variable section of object (all protected by properties)
Private lpOSInfo As OSVERSIONINFO
'combined Major/minor version in initialize
Private sngOSVersion As Single
'variable for instance count on the RAS DLL
Private hRASDLLInstance As Long

Private Sub Class_Initialize()
    
   'Have to get windows version for all RAS functions
   Dim lngErrNum As Long
   Dim lngbugfix As Long
    
   'Initialize classes to avoid potential problems
   'as documented in Q143261
   Set lpRASError = New RASError
   lngbugfix = lpRASError.ErrorNumber
   Set lpPhoneEntries = New PhoneEntries
   lngbugfix = lpPhoneEntries.Count
   Set lpConnections = New Connections
   lngbugfix = lpConnections.Count
   'Initialize RAS
   lngErrNum = fcnLoadandCheckRAS()
   If lngErrNum Then
      'Failed so raise error to fail
      Err.Raise vbObjectError + 1911, "RAS.RASEngine", "RAS Could Not Be Initialized. WIN32 Error: " & Str$(lngErrNum)
   Else
      'We have RAS on system and it is now running
      'Have to set the structure size
      lpOSInfo.dwOSVersionInfoSize = 148
      If (GetVersionEx(lpOSInfo)) Then
         'set the global version variable for use in all other RAS functions. This is decl in the BAS file.
         lngWindowVersion = lpOSInfo.dwPlatformId
         'Have to combine the two DWORDS into a Single (I will bet that there is a more efficient method, but...)
         sngOSVersion = CSng(lpOSInfo.dwMajorVersion) + CSng(Val("." & Str$(lpOSInfo.dwMinorVersion)))
      Else
         lngErrNum = GetLastError()
         'This will cause Initialization failure
         Err.Raise vbObjectError + 1912, "RAS.RASEngine", "GetVersionEx Failed With WIN32 Error: " & Str$(lngErrNum)
      End If
   End If
       
End Sub





Public Property Get OSVersion() As Single

   'Have to combine the two DWORDS into a Single
   OSVersion = sngOSVersion

End Property



Public Property Get OSBuildNumber() As Long
   
   OSBuildNumber = lpOSInfo.dwBuildNumber

End Property


Public Property Get OSType() As Long
   
   'WIN32s = 0
   'WIN 95 = 1
   'WIN NT = 2
   OSType = lpOSInfo.dwPlatformId

End Property



Public Function RASError() As RASError
   
   'set up a new pointer to Error object
   Set RASError = lpRASError
   
End Function



Public Function Connections(Optional ByVal Index As Variant) As Object

   If IsMissing(Index) Then
      'set pointer to my Connections collection object
      Set Connections = lpConnections
   Else
      'set pointer to the Connection Object
      Set Connections = arrConnection(Index)
   End If
 
End Function



Public Function PhoneEntries(Optional ByVal Index As Variant) As Object
   
   If IsMissing(Index) Then
      'set pointer to my PhoneEntries collection object
      Set PhoneEntries = lpPhoneEntries
   Else
      'set pointer to the PhoneEntry Object
      Set PhoneEntries = arrPEntry(Index)
   End If
   
End Function

Private Function fcnLoadandCheckRAS() As Long

   'Checks out RAS and loads DLLs at startup so faster later.
   hRASDLLInstance = LoadLibrary("RASAPI32.DLL")
   If hRASDLLInstance Then
      fcnLoadandCheckRAS = 0
   Else
      fcnLoadandCheckRAS = GetLastError()
   End If
    
End Function

Private Function fcnUnloadRAS() As Long

   Dim lngRetCode As Long
    
   'Just decrementing my instance count on RASAPI32.DLL from my LoadLibrary at startup.
   lngRetCode = FreeLibrary(hRASDLLInstance)
   If lngRetCode Then
      fcnUnloadRAS = 0
   Else
      fcnUnloadRAS = GetLastError()
   End If
    
End Function

Private Sub Class_Terminate()
   
   Dim lngErrNum As Long
   Dim intLooper As Long
   
   'Discarding errors because this is just in case user is being careless with connections.
   'This can be removed if want connections to remain open after server is terminated.
   On Error Resume Next
   For intLooper = 0 To lpConnections.Count - 1
      lpConnections.RemoveConnection (intLooper)
   Next intLooper
   On Error GoTo 0
   
   'Decrement the function instance handle.
   'as documented in Q143261
   Set lpPhoneEntries = Nothing
   Set lpConnections = Nothing
   Set lpRASError = Nothing
   lngErrNum = fcnUnloadRAS()
   'I do not expect to see this error, but if I do I want to let the user know
   If lngErrNum Then Err.Raise vbObjectError + 1913, "RAS.RASEngine", "RAS was not properly uninitialized. WIN32 Error: " & Str$(lngErrNum)
   
End Sub
