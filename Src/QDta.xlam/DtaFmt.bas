Attribute VB_Name = "DtaFmt"
'Option Explicit
'Function DrsLyInsBrkLin(TblLy$(), ColNm$) As String()
'Dim Hdr$: Hdr = TblLy(1)
'Dim Fny$():
'    Fny = SplitVBar(Hdr)
'    Fny = AyRmvFstEle(Fny)
'    Fny = AyRmvLasEle(Fny)
'    Fny = SyTrim(Fny)
'Dim Ix%
'    Ix = AyIx(Fny, ColNm)
'Dim DryLy$()
'    DryLy = AyWhExclAtCnt(TblLy, 0, 2)
'Dim O$()
'    Push O, TblLy(0)
'    Push O, TblLy(1)
'    PushAy O, DryLy_InsBrkLin(DryLy, Ix)
'DrsLyInsBrkLin = O
'End Function
'
'
'Sub Tst__VbFmt()
'DrsLyInsBrkLin__Tst
'End Sub
'
'Private Sub DrsLyInsBrkLin__Tst()
'Dim TblLy$()
'Dim Act$()
'Dim Exp$()
'TblLy = FtLy(TstResPth & "DrsLyInsBrkLin.txt")
'Act = DrsLyInsBrkLin(TblLy, "Tbl")
'Exp = FtLy(TstResPth & "DrsLyInsBrkLin_Exp.txt")
'AyPair_EqChk Exp, Act
'End Sub
'
'Function DaoTyToSim(T As DataTypeEnum) As eSimTy
'Dim O As eSimTy
'Select Case T
'Case _
'   DAO.DataTypeEnum.dbBigInt, _
'   DAO.DataTypeEnum.dbByte, _
'   DAO.DataTypeEnum.dbCurrency, _
'   DAO.DataTypeEnum.dbDecimal, _
'   DAO.DataTypeEnum.dbDouble, _
'   DAO.DataTypeEnum.dbFloat, _
'   DAO.DataTypeEnum.dbInteger, _
'   DAO.DataTypeEnum.dbLong, _
'   DAO.DataTypeEnum.dbNumeric, _
'   DAO.DataTypeEnum.dbSingle
'   O = eNbr
'Case _
'   DAO.DataTypeEnum.dbChar, _
'   DAO.DataTypeEnum.dbGUID, _
'   DAO.DataTypeEnum.dbMemo, _
'   DAO.DataTypeEnum.dbText
'   O = eTxt
'Case _
'   DAO.DataTypeEnum.dbBoolean
'   O = eLgc
'Case _
'   DAO.DataTypeEnum.dbDate, _
'   DAO.DataTypeEnum.dbTimeStamp, _
'   DAO.DataTypeEnum.dbTime
'   O = eDte
'Case Else
'   O = eOth
'End Select
'DaoTyToSim = O
'End Function
'
'Sub Fiy(Fny$(), FldLvs$, ParamArray OAp())
''Fiy=Field Index Array
'Dim A$(): A = SplitSpc(FldLvs)
'Dim I&(): I = AyIxAy(Fny, A)
'Dim J%
'For J = 0 To UB(I)
'    OAp(J) = I(J)
'Next
'End Sub
'
'Function IsSimTyLvs(A$) As Boolean
'Dim Ay$(): Ay = SslSy(A)
'If AyIsEmp(Ay) Then Exit Function
'Dim I
'For Each I In Ay
'   If Not IsSimTyStr(Ay) Then Exit Function
'Next
'IsSimTyLvs = True
'End Function
'
'Function IsSimTyStr(S) As Boolean
'Select Case UCase(S)
'Case "TXT", "NBR", "LGC", "DTE", "OTH": IsSimTyStr = True
'End Select
'End Function
'
'Function ItrCntByBoolPrp&(A, BoolPrpNm$)
'If A.Count = 0 Then Exit Function
'Dim O, Cnt&
'For Each O In A
'    If CallByName(O, BoolPrpNm, VbGet) Then
'        Cnt = Cnt + 1
'    End If
'Next
'ItrCntByBoolPrp = Cnt
'End Function
'
'Function ItrItmByPrp(A, PrpNm$, PrpV)
'Dim O, V
'If A.Count > 0 Then
'    For Each O In A
'        V = CallByName(O, PrpNm, VbGet)
'        If V = PrpV Then
'            Asg O, ItrItmByPrp
'            Exit Function
'        End If
'    Next
'End If
'End Function
'
'Function ItrNy(A, Optional Lik$ = "*") As String()
'Dim O$(), Obj, N$
'If A.Count > 0 Then
'    For Each Obj In A
'        N = Obj.Name
'        If N Like Lik Then Push O, N
'    Next
'End If
'ItrNy = O
'End Function
'


