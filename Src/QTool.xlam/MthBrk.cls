VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MthBrk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Nm$, Mdy$, Ty$

Friend Property Get Init(Nm$, Mdy$, Ty$) As MthBrk
Me.Nm = Nm
Me.Mdy = Mdy
Me.Ty = Ty
Ass MthTy_IsVdt(Ty)
Set Init = Me
End Property

Property Get ToStr$()
If Mdy = "" Then
    ToStr = FmtQQ("MthBrk(Nm(?) Ty(?))", Nm, Ty)
Else
    ToStr = FmtQQ("MthBrk(Nm(?) Mdy(?) Ty(?))", Nm, Mdy, Ty)
End If
End Property

