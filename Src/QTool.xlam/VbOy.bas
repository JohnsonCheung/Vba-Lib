Attribute VB_Name = "VbOy"
Option Explicit
Function OyPrpAy(Oy, PrpNm) As Variant()
OyPrpAy = OyPrpAyInto(Oy, PrpNm, EmpAy)
End Function
Function OyPrpAyInto(Oy, PrpNm, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(Oy) > 0 Then
    Dim I
    For Each I In Oy
        Push O, ObjPrp(I, PrpNm)
    Next
End If
OyPrpAyInto = O
End Function
Function OyNy(Oy) As String()
Dim O$(): If Sz(Oy) = 0 Then Exit Function
Dim I
For Each I In Oy
    Push O, CallByName(I, "Name", VbGet)
Next
OyNy = O
End Function
Function OyToStrSy(A) As String()
If Sz(A) = 0 Then Exit Function
Dim O$()
ReDim O(UB(A))
Dim J&
For J = 0 To UB(A)
    O(J) = A(J).ToStr
Next
OyToStrSy = O
End Function
Function OyWhPrpIn(A, P, InAy)
Dim X, O
If Sz(A) = 0 Or Sz(InAy) Then OyWhPrpIn = A: Exit Function
O = A
Erase O
For Each X In A
    If AyHas(InAy, ObjPrp(X, P)) Then PushObj O, X
Next
OyWhPrpIn = O
End Function
Function OyWhNmPatnExl(A, Patn$, ExlAy$)
OyWhNmPatnExl = OyWhNmExl(OyWhNmPatn(A, Patn), ExlAy)
End Function
Function OyWhNmExl(A, ExlAy$)
If ExlAy = "" Then OyWhNmExl = A: Exit Function
Dim X, LikAy$(), O
O = A
Erase O
LikAy = SslSy(ExlAy)
For Each X In A
    If Not IsInLikAy(X.Name, LikAy) Then PushObj O, X
Next
OyWhNmExl = O
End Function
Function OyWhNmPatn(A, Patn$)
If Patn = "." Then OyWhNmPatn = A: Exit Function
Dim X, O, Re As New RegExp
O = A
Erase O
Re.Pattern = Patn
For Each X In A
    If Re.Test(X.Name) Then PushObj O, X
Next
OyWhNmPatn = O
End Function
Function OyWhNm(A, Re As RegExp, ExlAy)
If Sz(A) = 0 Then OyWhNm = A: Exit Function
Dim X
For Each X In A
    If IsNmSel(X.Name, Re, ExlAy) Then PushObj OyWhNm, X
Next
End Function
Sub OyDo(Oy, DoFun$)
Dim O
For Each O In Oy
    Excel.Run DoFun, O ' DoFunNm cannot be like a Excel.Address (eg, A1, XX1)
Next
End Sub
