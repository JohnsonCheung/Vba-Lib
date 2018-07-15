Attribute VB_Name = "M_Dr"
Option Explicit

Property Get DrBySsl(Ssl$, TyAy() As eSimTy) As Variant()
Stop
End Property

Property Get DrLin$(Dr, Wdt%())
Dim UDr%
   UDr = UB(Dr)
Dim O$()
   Dim U1%: U1 = UB(Wdt)
   ReDim O(U1)
   Dim W, V
   Dim J%, V1$
   J = 0
   For Each W In Wdt
       If UDr >= J Then V = Dr(J) Else V = ""
       V1 = AlignL(V, W)
       O(J) = V1
       J = J + 1
   Next
DrLin = Quote(Join(O, " | "), "| * |")
End Property

Property Get DryBy_Ay_and_Const(Ay, Constant) As Variant()
If AyIsEmp(Ay) Then Exit Property
Dim O(), I
For Each I In Ay
   Push O, Array(I, Constant)
Next
DryBy_Ay_and_Const = O
End Property

Property Get DryBy_Const_and_Ay(Constant, Ay) As Variant()
If AyIsEmp(Ay) Then Exit Property
Dim O(), I
For Each I In Ay
   Push O, Array(Constant, I)
Next
DryBy_Const_and_Ay = O
End Property

Property Get DryColColl(A, ColIx%) As Collection
Dim O As New Collection
If Not Sz(A) = 0 Then
    Dim Dr
    For Each Dr In A
        O.Add Dr(ColIx)
    Next
End If
Set DryColColl = O
End Property

Property Get DryStrDry(A, ShwZer As Boolean) As Variant()
Dim O(), Dr
For Each Dr In A
   Push O, AyCellSy(Dr, ShwZer)
Next
DryStrDry = O
End Property

Property Get S1S2Ay_Dry(A() As S1S2) As Variant()
Dim O()
Dim J&
For J = 0 To UB(A)
   With A(J)
       Push O, Array(.S1, .S2)
   End With
Next
S1S2Ay_Dry = O
End Property

Property Get VblLy_Dry(A$()) As Variant()
Dim O(), I
If Sz(A) = 0 Then Exit Property
For Each I In A
    Push O, SplitVBar(I, Trim:=True)
Next
VblLy_Dry = O
End Property

Function DicDry(A As Dictionary, Optional InclDicValTy As Boolean) As Variant()
Dim O(), I
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Not AyIsEmp(K) Then
   If InclDicValTy Then
       For Each I In K
           Push O, Array(I, A(I), TypeName(A(I)))
       Next
   Else
       For Each I In K
           Push O, Array(I, A(I))
       Next
   End If
End If
DicDry = O
End Function

Function DrExpLinesCol(Dr, LinesColIx%) As Variant()
Dim B$()
    B = SplitCrLf(CStr(Dr(LinesColIx)))
Dim O()
    Dim IDr
        IDr = Dr
    Dim I
    For Each I In B
        IDr(LinesColIx) = I
        Push O, IDr
    Next
DrExpLinesCol = O
End Function

Private Property Get DrIx_IsBrk(Dr, DrIx&, BrkColIx%) As Boolean
If AyIsEmp(Dr) Then Exit Property
If DrIx = 0 Then Exit Property
If DrIx = UB(Dr) Then Exit Property
If Dr(DrIx)(BrkColIx) = Dr(DrIx - 1)(BrkColIx) Then Exit Property
DrIx_IsBrk = True
End Property

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
