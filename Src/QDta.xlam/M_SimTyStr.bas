Attribute VB_Name = "M_SimTyStr"
Option Explicit

Function SimTyStr_SimTy(A) As eSimTy
Dim O As eSimTy
Select Case UCase(A)
Case "TXT": O = eTxt
Case "NBR": O = eNbr
Case "LGC": O = eLgc
Case "DTE": O = eDte
Case Else: O = eOth
End Select
SimTyStr_SimTy = O
End Function
