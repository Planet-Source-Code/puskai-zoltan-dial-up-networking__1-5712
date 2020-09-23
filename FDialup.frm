VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "Comctl32.ocx"
Begin VB.Form FDialup 
   Caption         =   "Some Dialup Networking Function"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDialEntry 
      Caption         =   "Dial Entry"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdHagupEntry 
      Caption         =   "Hang Up Entry"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Timer tmrGetConnStatus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton cmdRenameEntry 
         Caption         =   "Rename Entry"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteEntry 
         Caption         =   "Delete Entry"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdEditEntry 
         Caption         =   "Edit Entry"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdCreateEntry 
         Caption         =   "Create Entry"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin ComctlLib.StatusBar statConnection 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   3315
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3245
            MinWidth        =   3245
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   14182
            MinWidth        =   14182
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstEntries 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblEntries 
      Caption         =   "D.U.N Entries :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FDialup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objRASConn As Connection
Private objRASEng As RASEngine
Private hRasConn As Long


Private Sub cmdCreateEntry_Click()
Dim retVal As Long
retVal = RasCreatePhonebookEntry(FDialup.hWND, vbNullString)
EnumEntries lstEntries
End Sub

Private Sub cmdDeleteEntry_Click()
Dim retVal As Long
If lstEntries.ListIndex >= 0 Then
    If vbYes = MsgBox("Are You sure, You want to delete '" & lstEntries.List(lstEntries.ListIndex) & "' Entry ?", vbYesNo + vbCritical, "Dialup") Then
        retVal = RasDeleteEntry(vbNullString, lstEntries.List(lstEntries.ListIndex))
    End If
End If
EnumEntries lstEntries
End Sub

Private Sub cmdDialEntry_Click()
Set objRASEng = Nothing
Set objRASEng = New RASEngine
If lstEntries.ListIndex >= 0 Then
    statConnection.Panels(1).Text = lstEntries.List(lstEntries.ListIndex)
    Set objRASConn = objRASEng.PhoneEntries(lstEntries.ListIndex).DialEntry(1)
    hRasConn = "&H" & FDialup.Tag
    tmrGetConnStatus.Enabled = True
End If
End Sub

Private Sub cmdEditEntry_Click()
Dim retVal As Long
If lstEntries.ListIndex >= 0 Then
    retVal = RasEditPhonebookEntry(FDialup.hWND, vbNullString, lstEntries.List(lstEntries.ListIndex))
End If
EnumEntries lstEntries
End Sub

Private Sub cmdHagupEntry_Click()
Dim retVal As Long
   
'Hang up the connection
tmrGetConnStatus.Enabled = False
retVal = RasHangUp(hRasConn)
statConnection.Panels(1).Text = ""
statConnection.Panels(2).Text = ""
End Sub

Private Sub cmdRenameEntry_Click()
Dim strNewEntryName As String
Dim retVal As Long
If lstEntries.ListIndex >= 0 Then
    strNewEntryName = InputBox("Type the  new Entry name for Entry: " & lstEntries.List(lstEntries.ListIndex), "Dialup")
    If RasValidateEntryName(vbNullString, strNewEntryName) = 0 Then
        retVal = RasRenameEntry(vbNullString, lstEntries.List(lstEntries.ListIndex), strNewEntryName)
    End If
End If
EnumEntries lstEntries
End Sub

Private Sub Form_Load()
    EnumEntries lstEntries
End Sub


Private Sub EnumEntries(lst As ListBox)

Dim s As Long, l As Long, ln As Long, a$
ReDim R(255) As RASENTRYNAME95
    
    lst.Clear
    R(0).dwSize = 264
    s = 256 * R(0).dwSize
    l = RasEnumEntries(vbNullString, vbNullString, R(0), s, ln)
    For l = 0 To ln - 1
        a$ = StrConv(R(l).szEntryName(), vbUnicode)
        lst.AddItem Left$(a$, InStr(a$, Chr$(0)) - 1)
    Next
    
    If lst.ListCount > 0 Then
        lst.ListIndex = 0
    End If

End Sub

Private Sub tmrGetConnStatus_Timer()

  Dim lngRetCode As Long
  Dim lngRASConnState As Long
  Dim lngRASError As Long
    
   If lngWindowVersion = 2 Then
      'using NT
      Dim lpRASCONNSTATUS As RASCONNSTATUS
      lpRASCONNSTATUS.dwSize = 64
      lngRetCode = RasGetConnectStatus(hRasConn, lpRASCONNSTATUS)
      If lngRetCode Then
         'Get and display error
         lngRASErrorNumber = lngRetCode
         'set status text
         statConnection.Panels(2).Text = lpRASError.fcnRASErrorString()
         lngRetCode = RasHangUp(hRasConn)
         'allow user to close
         MsgBox lpRASError.fcnRASErrorString(), vbOKOnly + vbCritical, "Dialup"
         statConnection.Panels(1).Text = ""
         statConnection.Panels(2).Text = ""
         'disable timer
         tmrGetConnStatus.Enabled = False
         lngRASError = 10000
      Else
         'success
         lngRASConnState = lpRASCONNSTATUS.rasconnstate
         lngRASError = lpRASCONNSTATUS.dwError
      End If
   Else
      'using 95
      Dim lpRASCONNSTATUS95 As RASCONNSTATUS95
      lpRASCONNSTATUS95.dwSize = 160
      lngRetCode = RasGetConnectStatus(hRasConn, lpRASCONNSTATUS95)
      If lngRetCode Then
         'Get and display error
         lngRASErrorNumber = lngRetCode
         'set status text
         statConnection.Panels(2).Text = lpRASError.fcnRASErrorString()
         lngRetCode = RasHangUp(hRasConn)
         'allow user to close
         cmdClose.Enabled = True
         'disable timer
         tmrGetConnStatus.Enabled = False
         lngRASError = 10000
      Else
         'success
         lngRASConnState = lpRASCONNSTATUS95.rasconnstate
         lngRASError = lpRASCONNSTATUS95.dwError
      End If
   End If
   
   'If Error then raise it else update the textbox with the appropriate info
   Select Case lngRASError
      Case SUCCESS, PENDING
         'Update connection
         Select Case lngRASConnState
            Case RASCS_OpenPort
               statConnection.Panels(2).Text = "Attempting To Open Port..."
            Case RASCS_PortOpened
               statConnection.Panels(2).Text = "Port Successfully Opened"
            Case RASCS_ConnectDevice
               statConnection.Panels(2).Text = "Attempting to Connect Device..."
            Case RASCS_DeviceConnected
               statConnection.Panels(2).Text = "Device Opened"
            Case RASCS_AllDevicesConnected
               statConnection.Panels(2).Text = "All Devices Connected"
            Case RASCS_Authenticate
               statConnection.Panels(2).Text = "Authenticating..."
            Case RASCS_AuthNotify
               statConnection.Panels(2).Text = "Athentication Notification"
            Case RASCS_AuthRetry
               statConnection.Panels(2).Text = "Retrying Authentication..."
            Case RASCS_AuthCallback
               statConnection.Panels(2).Text = "Authentication Callback"
            Case RASCS_AuthChangePassword
               statConnection.Panels(2).Text = "Change Password"
            Case RASCS_AuthProject
               statConnection.Panels(2).Text = "Authenticating Project.."
            Case RASCS_AuthLinkSpeed
               statConnection.Panels(2).Text = "Authenticating Link Speed.."
            Case RASCS_AuthAck
               statConnection.Panels(2).Text = "Athentication Acknowlegement"
            Case RASCS_ReAuthenticate
               statConnection.Panels(2).Text = "ReAuthentication..."
            Case RASCS_Authenticated
               statConnection.Panels(2).Text = "Authenticated"
            Case RASCS_PrepareForCallback
               statConnection.Panels(2).Text = "Prepare For Callback"
            Case RASCS_WaitForModemReset
               statConnection.Panels(2).Text = "Waiting For Modem Rest..."
            Case RASCS_WaitForCallback
               statConnection.Panels(2).Text = "Waiting For Callback..."
            Case RASCS_Projected
               statConnection.Panels(2).Text = "Network Completely Configured"
            Case RASCS_StartAuthentication    'Windows 95 only
               statConnection.Panels(2).Text = "Attempting to Open Port"
            Case RASCS_CallbackComplete         'Windows 95 only
               statConnection.Panels(2).Text = "Callback Completed"
            Case RASCS_LogonNetwork            'Windows 95 only
               statConnection.Panels(2).Text = "Logging On To Network"
            Case RASCS_Interactive
               statConnection.Panels(2).Text = "Interactive"
            Case RASCS_RetryAuthentication
               statConnection.Panels(2).Text = "Retry Authentication"
            Case RASCS_CallbackSetByCaller
               statConnection.Panels(2).Text = "CallBack Set By Caller"
            Case RASCS_PasswordExpired
               statConnection.Panels(2).Text = "Password Expired"
            Case RASCS_Connected
               statConnection.Panels(2).Text = "Connected"
               cmdClose.Enabled = True
            Case RASCS_Disconnected
               statConnection.Panels(2).Text = "Disconnected"
            Case Else
               statConnection.Panels(2).Text = "Unknown State"
         End Select
      Case 10000
         'do nothing because RasGetConnectStatus failed
      Case Else
         'We have an error
         lngRASErrorNumber = lngRASError
         'set status text
         statConnection.Panels(2).Text = lpRASError.fcnRASErrorString()
         'Hang up the connection
         lngRetCode = RasHangUp(hRasConn)
         'allow user to close
         cmdClose.Enabled = True
         'disable timer
         tmrGetConnStatus.Enabled = False
   End Select
   
End Sub
