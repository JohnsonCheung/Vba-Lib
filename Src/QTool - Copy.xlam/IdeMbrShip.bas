Attribute VB_Name = "IdeMbrShip"
Option Explicit
Function PjHasMd(A As VBProject, Nm) As Boolean
Dim T As vbext_ComponentType
If Not ItrHasNm(A.VBComponents, Nm) Then Exit Function
T = PjCmp(A, Nm).Type
If T = vbext_ct_StdModule Then PjHasMd = True: Exit Function
Debug.Print "PjHasMd: Pj(?) has Mbr(?), but it is not Md, but CmpTy(?)", PjNm(A), Nm, CmpTyStr(T)
End Function

