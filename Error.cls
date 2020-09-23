VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RASError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declarations for Error object
'************************************************************************************************************************************************************************************

Public Function fcnRASErrorString() As String
    
   Dim lngRetCode As Long
   'Error String Buffer. 256 byte is large enough for all errors
   Dim strRASErrorString As String
    
   'Fill string
   strRASErrorString = Space$(256)
   'lngRASErrorNumber is the RAS error number in class decl
   Select Case lngRASErrorNumber
      Case Is >= 600
         lngRetCode = RasGetErrorString(lngRASErrorNumber, strRASErrorString, 256&)
         If lngRetCode Then
            'We should never see this
            fcnRASErrorString = "Call To RasGetErrorString Failed. Error Retrieving The Error String!!!"
         Else
            'Return string
            fcnRASErrorString = strRASErrorString
         End If
      Case NOT_SUPPORTED
         fcnRASErrorString = "Function Is Not Supported On This Version of Windows."
      Case 6
         fcnRASErrorString = "Megszakadt a telefonvonal!"
      Case Else
         fcnRASErrorString = "Unexpected Error"
   End Select
    
End Function
Public Property Get ErrorNumber() As Long

   'Property seems trivial, but I may need to expand in future.
   ErrorNumber = lngRASErrorNumber

End Property



Public Property Get Description() As String
   
   'Property seems trivial, but I may need to expand in future.
   Description = strRASDescription

End Property

Public Sub Clear()

   'Gets rid of the errors. For use in error handling routine of Client.
   lngRASErrorNumber = 0
   strRASDescription = ""
   
End Sub

