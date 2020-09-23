VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhoneEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'variables for entry object. I could use a type but then I have to have 2
'and StrConv on all property gets
Private strEntryName As String
Private strPhoneNumber As String
Private strCallbackNumber As String
Private strUserName As String
Private strPassword As String
Private strDomain As String
Private intIndex As Integer
'lets me know if Password was fetched
Private lngGotPassword As Long

Public Property Get EntryName() As String

   EntryName = strEntryName

End Property

Public Property Let EntryName(strNewName As String)

   Dim lngRetCode As Long

   'Test flag to be sure that it is not client
   If boolAllowUpdate Then
      strEntryName = strNewName
      lngRetCode = fcnRasGetEntryDialParams()
      If lngRetCode Then
         Err.Raise vbObjectError + lngRetCode, "EntryName Set Failed", "RAS Failure"
      End If
   Else
      lngRASErrorNumber = 1111
      strRASDescription = "Property Not Updateable"
      Err.Raise vbObjectError + 1111, "Property Not Updateable", "RAS Failure"
   End If

End Property

Public Function DialEntry(boolAsync As Boolean) As Connection

   'add a connection
   Set DialEntry = lpConnections.AddConnection(strEntryName, "", "", strUserName, strPassword, "", boolAsync)

End Function

Public Sub EditEntry()

   Dim lngRetCode As Long
   
   'RAS calls up dialog under 95. Fails on NT
   lngRetCode = RasEditPhonebookEntry(APINULL, vbNullString, strEntryName)
    If lngRetCode Then
      lngRASErrorNumber = lngRetCode
      strRASDescription = lpRASError.fcnRASErrorString()
      Err.Raise vbObjectError + lngRetCode, "EditEntry Method Failed", "RAS Failure"
    End If

End Sub

Private Function fcnRasGetEntryDialParams() As Long

   Dim lngRetCode As Long
   Dim lngRetlstrcpy As Long
   'not bothering with the NT case because I know it fails.
   Dim lprasdialparams As RASDIALPARAMS95
   
   'set up the structure
   lprasdialparams.dwSize = 1052
   lngRetlstrcpy = lstrcpy(lprasdialparams.szEntryName(0), strEntryName)
   lngRetCode = RasGetEntryDialParams(vbNullString, lprasdialparams, lngGotPassword)
   Select Case lngRetCode
      Case SUCCESS
         strUserName = fcnTrimNulls(StrConv(lprasdialparams.szUserName, vbUnicode))
         strPhoneNumber = fcnTrimNulls(StrConv(lprasdialparams.szPhoneNumber, vbUnicode))
         strCallbackNumber = fcnTrimNulls(StrConv(lprasdialparams.szCallbackNumber, vbUnicode))
         If lngGotPassword = 1 Then
            strPassword = fcnTrimNulls(StrConv(lprasdialparams.szPassword, vbUnicode))
         Else
            strPassword = "Password Not Available"
         End If
         strDomain = fcnTrimNulls(StrConv(lprasdialparams.szDomain, vbUnicode))
         fcnRasGetEntryDialParams = 0
      Case NOT_SUPPORTED
         'running NT so set to "Not Available"
         strUserName = "Not Available"
         strPhoneNumber = "Not Available"
         strCallbackNumber = "Not Available"
         strPassword = "Not Available"
         strDomain = "Not Available"
         'Not going to fail on NT because this is extraneous info
         fcnRasGetEntryDialParams = 0
      Case Else
         strUserName = "Not Available"
         strPhoneNumber = "Not Available"
         strCallbackNumber = "Not Available"
         strPassword = "Not Available"
         strDomain = "Not Available"
         lngRASErrorNumber = lngRetCode
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         fcnRasGetEntryDialParams = lngRetCode
   End Select
   
End Function

Private Function fcnRasSetEntryDialParams() As Long

   Dim lngRetCode As Long
   Dim lngRetlstrcpy As Long
   Dim lngAttemptOrder As Long
    'not bothering with the NT case because I know it fails.
   Dim lprasdialparams As RASDIALPARAMS95
    
   lprasdialparams.dwSize = 1052
   'I should trap, but I'll catch it later
   lngRetlstrcpy = lstrcpy(lprasdialparams.szEntryName(0), strEntryName)
   lngRetlstrcpy = lstrcpy(lprasdialparams.szUserName(0), strUserName)
   lngRetlstrcpy = lstrcpy(lprasdialparams.szDomain(0), strDomain)
   lngRetlstrcpy = lstrcpy(lprasdialparams.szPhoneNumber(0), strPhoneNumber)
   lngRetlstrcpy = lstrcpy(lprasdialparams.szCallbackNumber(0), strCallbackNumber)
   'do not want garbage as password by mistake
   If lngGotPassword = 1 Then
      lngRetlstrcpy = lstrcpy(lprasdialparams.szPassword(0), strPassword)
   Else
      lngRetlstrcpy = lstrcpy(lprasdialparams.szPassword(0), "")
   End If
   lngRetCode = RasSetEntryDialParams(vbNullString, lprasdialparams, 0&)
   Select Case lngRetCode
      Case SUCCESS
         strUserName = fcnTrimNulls(StrConv(lprasdialparams.szUserName, vbUnicode))
         strPhoneNumber = fcnTrimNulls(StrConv(lprasdialparams.szPhoneNumber, vbUnicode))
         strCallbackNumber = fcnTrimNulls(StrConv(lprasdialparams.szCallbackNumber, vbUnicode))
         If lngGotPassword Then
            strPassword = fcnTrimNulls(StrConv(lprasdialparams.szPassword, vbUnicode))
         Else
            strPassword = "Not Available"
         End If
         strDomain = fcnTrimNulls(StrConv(lprasdialparams.szDomain, vbUnicode))
         fcnRasSetEntryDialParams = 0
      Case NOT_SUPPORTED
         'running NT so set to "Not Available"
         strUserName = "Not Available"
         strPhoneNumber = "Not Available"
         strCallbackNumber = "Not Available"
         strPassword = "Not Available"
         strDomain = "Not Available"
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         fcnRasSetEntryDialParams = lngRetCode
      Case Else
         'try to put the properties back in order
         lngAttemptOrder = fcnRasGetEntryDialParams()
         lngRASErrorNumber = lngRetCode
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         fcnRasSetEntryDialParams = lngRetCode
         
   End Select
    
End Function

Public Property Get PhoneNumber() As String

   PhoneNumber = strPhoneNumber

End Property

Public Property Let PhoneNumber(strNewNumber As String)

   Dim lngRetCode As Long
   
   strPhoneNumber = strNewNumber
   lngRetCode = fcnRasSetEntryDialParams()
   If lngRetCode Then
      Err.Raise vbObjectError + lngRetCode, "PhoneNumber Set Failed", "RAS Failure"
   End If
   
End Property

Public Property Get CallbackNumber() As String

   CallbackNumber = strCallbackNumber

End Property

Public Property Let CallbackNumber(strNewNumber As String)

   Dim lngRetCode As Long
   
   'set string. If this fails it will be set to 'Not Supported' because I have no
   'idea what went wrong
   strCallbackNumber = strNewNumber
   lngRetCode = fcnRasSetEntryDialParams()
   If lngRetCode Then
      Err.Raise vbObjectError + lngRetCode, "CallbackNumber Set Failed", "RAS Failure"
   End If

End Property

Public Property Get UserName() As String

   UserName = strUserName

End Property

Public Property Let UserName(strNewUser As String)

   Dim lngRetCode As Long
   
   strUserName = strNewUser
   lngRetCode = fcnRasSetEntryDialParams()
   If lngRetCode Then
      Err.Raise vbObjectError + lngRetCode, "UserName Set Failed", "RAS Failure"
   End If

End Property

Public Property Get Password() As String

   Password = strPassword

End Property

Public Property Let Password(strNewPassword As String)

   Dim lngRetCode As Long
   
   strPassword = strNewPassword
   lngRetCode = fcnRasSetEntryDialParams()
   If lngRetCode Then
      Err.Raise vbObjectError + lngRetCode, "Password Set Failed", "RAS Failure"
   End If

End Property

Public Property Get Domain() As String

   Domain = strDomain

End Property

Public Property Let Domain(strNewDomain As String)

   Dim lngRetCode As Long
   
   strDomain = strNewDomain
   lngRetCode = fcnRasSetEntryDialParams()
   If lngRetCode Then
      Err.Raise vbObjectError + lngRetCode, "Domain Set Failed", "RAS Failure"
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

