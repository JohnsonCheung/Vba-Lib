VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_Ay() As Blk
Friend Property Get Init(A() As Blk) As Blks
B_Ay = A
Set Init = Me
End Property

Function BlkAy_LnxAy(A() As Blk, BlkTyStr$) As Lnx()
Dim J%
For J = 0 To U
    If A(J).BlkTyStr = BlkTyStr Then LnxAy = A(J).LnxAy: Exit Function
Next
End Function
Property Get StmtsRslt() As StmtsRslt

End Property
Property Get BlkAy_ErLy(A() As Blk) As String()

End Property
Property Get BlkAy_SwRslt(A() As Blk) As SwRslt

End Property
Property Get BlkAy_PmRslt(A() As Blk) As PmRslt

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

Property Get BlkAy_SqLnxAy(A() As Blk) As Lnx()
Dim J%, O() As Lnxs
For J = 0 To UB(B_Ay)
    If B_Ay(J).BlkTyStr = "SQ" Then
        PushObj O, B_Ay(J).Lnxs
    End If
Next
SqLnxsAy = O
End Property

