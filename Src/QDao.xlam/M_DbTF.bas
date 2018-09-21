Attribute VB_Name = "M_DbTF"
Option Explicit
Function DbTF_Fld(A As Database, T, F) As DAO.Field
Set DbTF_Fld = A.TableDefs(T).Fields(F)
End Function
Function DbTF_FldInfDr(A As Database, T, F) As Variant()
Dim FF  As DAO.Field
Set FF = A.TableDefs(T).Fields(F)
With FF
    DbTF_FldInfDr = Array(F, IIf(DbTF_IsPk(A, T, F), "*", ""), DaoTy_Str(.Type), .Size, .DefaultValue, .Required, FldDes(FF))
End With
End Function
Function DbTF_IsPk(A As Database, T, F) As Boolean
DbTF_IsPk = AyHas(DbtPk(A, T), F)
End Function
Function DbTF_NxtId&(A As Database, T, Optional F)
Dim S$: S = FmtQQ("select Max(?) from ?", Dft(F, T), T)
DbTF_NxtId = DbqV(A, S) + 1
End Function
