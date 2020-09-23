VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhoneEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************************************************************************************************************************************************************************************


'Protected Variables for PhoneEntries Collection
'PhoneEnty object Count for Count Property
Private intPCount As Integer

Public Property Get Count() As Integer

   'Seems trivial, but I may need to expand later
   Dim lngRetCode

   'If this fails then count will be off
   lngRetCode = fcnRASEnumEntries()
   If lngRetCode Then
      Err.Raise vbObjectError + lngRetCode, "Connections.Count Failure", "RAS Failure"
   Else
      Count = intPCount
   End If

End Property

Private Function fcnRASEnumEntries() As Long
 
   Dim lngRetCode As Long
   Dim lpszreserved As String
   Dim lpszPhonebook As String
   Dim lpcb As Long
   Dim lpcEntries As Long
   Dim intArraySize As Integer
   Dim intLooper As Long
   
   lpszreserved = vbNullString
   lpszPhonebook = vbNullString
   'Putting a maximum of 256 Entries. If it fails then we resize
   intArraySize = 255
   If lngWindowVersion = 2 Then
      'We are running NT
      ReDim lprasentryname(intArraySize) As RASENTRYNAME
      lprasentryname(0).dwSize = 28
      lpcb = 256 * lprasentryname(0).dwSize
      lngRetCode = RasEnumEntries(lpszreserved, lpszPhonebook, lprasentryname(0), lpcb, lpcEntries)
   Else
      'We are running 95 (lngWindowVersion =1)
      ReDim lprasentryname95(intArraySize) As RASENTRYNAME95
      lprasentryname95(0).dwSize = 264
      lpcb = 256 * lprasentryname95(0).dwSize
      lngRetCode = RasEnumEntries(lpszreserved, lpszPhonebook, lprasentryname95(0), lpcb, lpcEntries)
   End If
   Select Case lngRetCode
      Case SUCCESS
         If lpcEntries > 0 Then
            'resize array so that it is correct size based on return from function
            ReDim arrPEntry(lpcEntries - 1) As PhoneEntry
            If lngWindowVersion = 2 Then
                'running NT
               For intLooper = 0 To UBound(arrPEntry())
                  Set arrPEntry(intLooper) = New PhoneEntry
                  'allow entryname update
                  boolAllowUpdate = True
                  arrPEntry(intLooper).EntryName = fcnTrimNulls(StrConv(lprasentryname(intLooper).szEntryName, vbUnicode))
                  arrPEntry(intLooper).Index = intLooper
                  boolAllowUpdate = False
               Next intLooper
            Else
               'running 95
               For intLooper = 0 To UBound(arrPEntry())
                  Set arrPEntry(intLooper) = New PhoneEntry
                  'allow entryname update
                  boolAllowUpdate = True
                  arrPEntry(intLooper).EntryName = fcnTrimNulls(StrConv(lprasentryname95(intLooper).szEntryName, vbUnicode))
                  arrPEntry(intLooper).Index = intLooper
                  boolAllowUpdate = False
               Next intLooper
            End If
         End If
         'set the Phoneentries.Count
         'I doubt that this will fail, but...
         intPCount = CInt(lpcEntries)
         fcnRASEnumEntries = 0
      Case ERROR_BUFFER_TOO_SMALL
         'Make buffers bigger and try again
         If lngWindowVersion = 2 Then
            'running NT
            intArraySize = lpcb / lprasentryname(0).dwSize
         Else
            'running 95
            intArraySize = lpcb / lprasentryname95(0).dwSize
         End If
      Case Else
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         fcnRASEnumEntries = lngRetCode
   End Select
  
End Function

Private Sub Class_Initialize()
   
   Dim lngSuccess As Long
   
   'initialize Phone book collection
   'throw away errors we will catch them later
   lngSuccess = fcnRASEnumEntries()
   
End Sub

Public Sub AddEntry()

   Dim lngRetCode As Long
   
   'Rather unconventional Add, but that is the way ras works
   'No easy way to return a guaranteed handle so I do not
   lngRetCode = RasCreatePhonebookEntry(APINULL, "")
      If lngRetCode Then
         lngRASErrorNumber = lngRetCode
         strRASDescription = lpRASError.fcnRASErrorString()
         Err.Raise vbObjectError + lngRetCode, "PhoneEntries AddEntry Failed", "RAS Failure"
      Else
         'no way to return handle to phoneentry object
         lngRetCode = fcnRASEnumEntries()
         If lngRetCode Then Err.Raise vbObjectError + lngRetCode, "PhoneEntries AddEntry Failed", "RAS Failure"
      End If

End Sub

