VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dbx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A As Database
Property Get Ds(Tny0) As Ds
Dim Tny$(): Tny = DftNy(Tny)
End Property

Private Sub ZZ_Qny()
AyDmp DbQny(CurDb)
End Sub

Property Get Inf() As DbInf
Set Inf = New DbInf
Stop
End Property

