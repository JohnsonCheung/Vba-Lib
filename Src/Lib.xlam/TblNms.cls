VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TblNms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Ny$()
Friend Property Get Init(Ny0) As TblNms
B_Ny = DftNy(Ny0)
Set Init = Me
End Property
Property Get DrpStmts$()
DrpStmts = JnDblCrLf(Oy.Sy("DrpStmt"))
End Property
Property Get Ny() As String()
Ny = B_Ny
End Property
Property Get N&()
N = Sz(B_Ny)
End Property
Property Get U&()
U = N - 1
End Property
Property Get TblNmAy() As TblNm()
Dim O() As TblNm, J%
For J = 0 To U
    PushObj O, TblNm(B_Ny(J))
Next
TblNmAy = O
End Property
Private Function Oy() As Oy
Dim O As New Oy
Set Oy = O.Init(TblNmAy)
End Function
