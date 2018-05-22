VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Md"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private A As CodeModule
Function Init(Md As CodeModule) As Md
Set A = Md
Set Init = Me
End Function
Function Nm$()

End Function
Function Val() As CodeModule
Set Val = A
End Function
