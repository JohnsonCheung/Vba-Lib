Attribute VB_Name = "M_Sample"
Option Explicit
Public Const SampleFx_KE24 = "C:\Users\User\Desktop\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"

Property Get SampleWs() As Worksheet
Dim O As Worksheet
Set O = NewWs
Stop
'DrsLo SampleDrs, WsRC(O, 2, 2)
Set SampleWs = O
WsVis O
End Property
Property Get SampleLo() As ListObject
Dim Ws As Worksheet: Set Ws = SampleWs
Set SampleLo = Ws.ListObjects(1)
End Property

Property Get SampleSq()
Dim O()
ReDim O(1 To 10, 1 To 7)
Dim J%, I%
For J = 1 To 7
    For I = 1 To 10
        O(I, J) = I * 10 + J
    Next
Next
SampleSq = O
End Property

