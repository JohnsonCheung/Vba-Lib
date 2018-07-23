Attribute VB_Name = "M_SimTyAy"
Option Explicit
Function SimTyAy_InsValTp$(SimTyAy() As eSimTy)
Dim U%
   U = UB(SimTyAy)
Dim Ay$()
   ReDim Ay(U)
Dim J%
For J = 0 To U
   Ay(J) = SimTy_QuoteTp(SimTyAy(J))
Next
SimTyAy_InsValTp = JnComma(Ay)
End Function
