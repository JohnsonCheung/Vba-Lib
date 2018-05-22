VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public IsVF As Boolean
Public V$
Public F$ 'If IsVF then F is FldLvs else F is Fld
Property Get Lin$()
If IsVF Then
    Lin = V & " " & F
Else
    Lin = F & " " & V
End If
End Property
