Attribute VB_Name = "M_DryLy"
Option Explicit

Function DryLy_InsBrkLin(DryLy$(), ColIx%) As String()
If Sz(DryLy) = 2 Then DryLy_InsBrkLin = DryLy: Exit Function
Dim Hdr$: Hdr = DryLy(0)
Dim Fm&, L%
   Dim N%: N = ColIx + 1
   Dim P1&, P2&
   P1 = InstrN(Hdr, "|", N)
   P2 = InStr(P1 + 1, Hdr, "|")
   Fm = P1 + 1
   L = P2 - P1 - 1
Dim O$()
   Push O, DryLy(0)
   Dim LasV$: LasV = Mid(DryLy(1), Fm, L)
   Dim J&
   Dim V$
   For J = 1 To UB(DryLy) - 1
       V = Mid(DryLy(J), Fm, L)
       If LasV <> V Then
           Push O, Hdr
           LasV = V
       End If
       Push O, DryLy(J)
   Next
   Push O, AyLasEle(DryLy)
DryLy_InsBrkLin = O
End Function
