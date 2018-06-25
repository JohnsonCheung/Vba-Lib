VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blk1s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_Ay() As Blk1
Friend Property Get Init(A() As Blk1) As Blk1s
B_Ay = A
Set Init = Me
End Property
Property Get N&()
N = Sz(B_Ay)
End Property
Property Get U&()
U = N - 1
End Property
Function LnxAy(BlkTyStr$) As Lnx1()
Dim J%
For J = 0 To U
    If B_Ay(J).BlkTyStr = BlkTyStr Then LnxAy = B_Ay(J).Lnxs.LnxAy: Exit Function
Next
End Function
Property Get StmtsRslt() As StmtsRslt

End Property
Property Get ErLy() As String()

End Property
Property Get SwRslt() As SwRslt

End Property
Property Get PmRslt() As PmRslt

End Property
Property Get ErBlkErLy() As String()
'?
End Property

Property Get ExcessPmBlkErLy() As String()
'?
End Property

Property Get ExcessSwBlkErLy() As String()
'?
End Property

Property Get SqLnxsAy() As Lnx1s()
Dim J%, O() As Lnx1s
For J = 0 To UB(B_Ay)
    If B_Ay(J).BlkTyStr = "SQ" Then
        PushObj O, B_Ay(J).Lnxs
    End If
Next
SqLnxsAy = O
End Property

