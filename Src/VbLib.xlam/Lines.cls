VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Lines$
Friend Property Get Init(Lines) As Lines
Me.Lines = Lines
End Property
Property Get EndTrim$()
EndTrim = JnCrLf(Ly(SplitCrLf(Ay)).EndTrim)
End Property

Private Sub ZZ_EndTrim()
Lines = RplVBar("lksdf|lsdfj|||")
Dim Act$: Act = EndTrim
Debug.Print Act & "<"
Stop
End Sub


