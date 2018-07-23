Attribute VB_Name = "M_VblLy"
Option Explicit

Function VblLy_Dry(VblLy$()) As Variant()
If AyIsEmp(VblLy) Then Exit Function
Dim O()
   Dim I
   For Each I In VblLy
       Push O, AyTrim(SplitVBar(CStr(I)))
   Next
VblLy_Dry = O
End Function

Sub VblLy_Dry__Tst()
Dim VblLy$()
Dim Exp$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
Dim Act()
Act = VblLy_Dry(VblLy)
DryBrw Act
End Sub

Private Sub ZZ_VblLy_Dry()
Dim VblLy$()
Dim Exp$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
Dim Act()
Act = VblLy_Dry(VblLy)
DryBrw Act
End Sub
