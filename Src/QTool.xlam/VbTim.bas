Attribute VB_Name = "VbTim"
Option Explicit
Sub TimFun(FunNy0)
Dim B!, E!, F
For Each F In DftNy(FunNy0)
    B = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - B
Next
End Sub
