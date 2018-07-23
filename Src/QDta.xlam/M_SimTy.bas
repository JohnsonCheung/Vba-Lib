Attribute VB_Name = "M_SimTy"
Option Explicit
Enum eSimTy
    eTxt
    eNbr
    eLgc
    eDte
    eOth
End Enum

Function SimTy_QuoteTp$(A As eSimTy)
Const CSub$ = "SimTyQuoteTp"
Dim O$
Select Case A
Case eTxt: O = "'?'"
Case eNbr, eLgc: O = "?"
Case eDte: O = "#?#"
Case Else
   Er CSub, "Given {eSimTy} should be [eTxt eNbr eDte eLgc]", A
End Select
SimTy_QuoteTp = O
End Function
