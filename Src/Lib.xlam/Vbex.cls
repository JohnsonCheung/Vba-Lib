VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vbex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A As Vbe
Friend Property Get Init(Vbe As Vbe) As Vbex
Set A = Vbe
Set Init = Me
End Property
Sub Export()
Dim I
For Each I In PjAy
    Pjx(I).Export
Next
End Sub
Property Get PjAy() As Pjx()
Dim O() As Pjx, I, M As Pjx
For Each I In A.VBProjects
    Set M = Pjx(I)
    If M.IsUsrLib Then
        PushObj O, M.Pj
    End If
Next
PjAy = O
End Property
Property Get PjNy() As String()
PjNy = Oy(PjAy).Ny
End Property
