VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Acs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public X As Access.Application
Property Get V() As Access.Application
Dim A$
On Error GoTo Er
A = X.Name
Set V = X
Exit Function
Er:
Set X = New Access.Application
Set V = X
End Property

Private Sub Class_Initialize()
Set X = New Access.Application
End Sub

Private Sub Class_Terminate()
On Error Resume Next
X.CloseCurrentDatabase
X.Quit
Set X = Nothing
End Sub
