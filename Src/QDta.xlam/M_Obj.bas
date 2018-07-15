Attribute VB_Name = "M_Obj"
Option Explicit

Function ObjPrpDr(Obj, PrpNy0) As Variant()
Dim Ny$(): Ny = DftNy(PrpNy0)
Dim U%
    U = UB(Ny)
Dim O()
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = CallByName(Obj, Ny(J), VbGet)
    Next
ObjPrpDr = O
End Function
