Attribute VB_Name = "JDta"
Option Explicit

Property Get Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(Fny0, Dry)
End Property

Property Get DrsByDic() As Drs
Stop
End Property

Property Get Ds(A() As Dt, Optional DsNm$ = "Ds") As Ds
Dim O As New Ds
Set Ds = O.Init(A, DsNm)
End Property

Property Get Dt(DtNm$, Fny0, Dry()) As Dt
Dim O As New Dt
Set Dt = O.Init(DtNm, Fny0, Dry)
End Property

Property Get SimTyStr_SimTy(A) As eSimTy
Dim O As eSimTy
Select Case UCase(A)
Case "TXT": O = eTxt
Case "NBR": O = eNbr
Case "LGC": O = eLgc
Case "DTE": O = eDte
Case Else: O = eOth
End Select
SimTyStr_SimTy = O
End Property

Sub SetPush(A As Dictionary, K)
If A.Exists(K) Then Exit Sub
A.Add K, Empty
End Sub

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
