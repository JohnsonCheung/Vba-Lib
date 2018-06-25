VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_Ay() As Boolean
Enum e_BoolAyOp
    e_And = 1
    e_Or = 2
    e_IsAllTrue = 3
    e_IsAllFalse = 4
    e_IsSomTrue = 5
    e_IsSomFalse = 6
End Enum
Sub Push(A As Boolean)
VbAy.Push B_Ay, A
End Sub
Property Get AndVal() As Boolean
AndVal = IsAllTrue
End Property
Property Get OrVal() As Boolean
OrVal = IsSomTrue
End Property
Function Val(Op As e_BoolAyOp) As Boolean
Select Case Op
Case e_And, e_IsAllTrue: Val = IsAllTrue
Case e_Or, e_IsSomTrue: Val = IsSomTrue
Case e_IsAllFalse: Val = IsAllFalse
Case e_IsSomFalse: Val = IsSomFalse
Case Else: Stop
End Select
End Function
Friend Property Get Init(A() As Boolean) As Bools
B_Ay = A
Set Init = Me
End Property
Property Get IsAllTrue() As Boolean
Dim J%
For J = 0 To UB(B_Ay)
    If Not B_Ay(J) Then Exit Property
Next
IsAllTrue = True
End Property
Property Get IsSomTrue() As Boolean
Dim J%
For J = 0 To UB(B_Ay)
    If B_Ay(J) Then IsSomTrue = True: Exit Property
Next
End Property
Property Get IsAllFalse() As Boolean
Dim J%
For J = 0 To UB(B_Ay)
    If B_Ay(J) Then Exit Property
Next
IsAllFalse = True
End Property
Property Get IsSomFalse() As Boolean
Dim J%
For J = 0 To UB(B_Ay)
    If Not B_Ay(J) Then IsSomFalse = True: Exit Property
Next
End Property

