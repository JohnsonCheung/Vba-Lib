Attribute VB_Name = "F_TstLoFmtr"
Option Explicit

Sub Tst(A As Range)
Dim Ws As Worksheet:
Dim Rg As Range
               Set Ws = RgWs(A)
Dim Tit$:         Tit = "Ix InpLoFmtrLy"
Dim TitAy$():   TitAy = SslSy(Tit)
                        AyRgH TitAy, WsRC(Ws, 2, 1) '<== Put Tit
                        WsA1(Ws).Value = "Msg"      '<== Put Msg Tit
Dim MsgRg As Range:
            Set MsgRg = WsRC(Ws, 1, 2)
                        MsgRg.Value = ""           '<== Clear Msg
Dim InpLyRg As Range:
               Set Rg = WsRC(Ws, 3, 4)
          Set InpLyRg = CellVBar(Rg)
Dim IsInRg As Boolean:
               IsInRg = CellIsInRgAp(A, InpLyRg)
                        If Not IsInRg Then
                            MsgRg.Value = "Not in range"
                            Exit Sub
                        End If
                                                          '<== ShwMsg not in range
Dim InpLy$():    InpLy = VBarSy(InpLyRg)
                         If InpLy(0) = "" Then
                            MsgRg.Value = "1st element of InpLy cannot be empty"
                            Exit Sub
                         End If                           '<== ShwMsg if no Input

Dim LnoRg As Range:
             Set LnoRg = RgRC(InpLyRg, 1, 0)
                         CellClrDown LnoRg
                         CellFillSeqDown LnoRg, Sz(InpLy)    '<== Put Lno: 0..
Stop
'Dim RsltDs As Ds
'                RsltDs = ZOupDs(InpLy)
Dim OupRg As Range
                Set Rg = WsRC(Ws, 3, 4)
             Set OupRg = CellVBar(Rg)
                         OupRg.Clear
                         'AyRgV(RsltLy, OupRg).Font = "Courier New"             '<== Put Rslt
End Sub

Private Sub ZZZ_Tstr_TotRslt(Rg As Range)
Stop '
'Static IsInChg As Boolean
'If IsInChg Then Exit Sub
'IsInChg = True
'Dim Ay(), J%, T$
''---------------
'Dim Ws As Worksheet:
'               Set Ws = RgWs(Rg)
'Dim Tit$:         Tit = "Lx *Tot Fld.. Fny Oup"
'Dim TitAy$():   TitAy = SslSy(Tit)
'                        AyRgH TitAy, WsRC(Ws, 2, 1) '<== Put Tit
'                        WsA1(Ws).Value = "Msg"      '<== Put Msg Tit
'Dim MsgRg As Range:
'            Set MsgRg = WsRC(Ws, 1, 2)
'                        MsgRg.Value = ""           '<== Clear Msg
'Dim InpLxRg As Range
'Dim InpTotRg As Range
'Dim InpFldLvsRg As Range
'Dim InpFnyRg As Range
'Dim OupRg As Range
'          Set InpLxRg = CellVBar(WsRC(Ws, 3, 1), AtLeastOneCell:=True)
'         Set InpTotRg = CellVBar(WsRC(Ws, 3, 2), AtLeastOneCell:=True)
'      Set InpFldLvsRg = CellVBar(WsRC(Ws, 3, 3), AtLeastOneCell:=True)
'         Set InpFnyRg = CellVBar(WsRC(Ws, 3, 4), AtLeastOneCell:=True)
'            Set OupRg = CellVBar(WsRC(Ws, 3, 5), AtLeastOneCell:=True)
'
'Dim IsInRg As Boolean:
'               IsInRg = CellIsInRgAp(Rg, InpTotRg, InpFldLvsRg)
'                        If Not IsInRg Then
'                            MsgRg.Value = "Not in range"
'                            GoTo X
'                        End If
'                                                   '<== ShwMsg not in range
'Dim InpTot$(): InpTot = VBarSy(InpTotRg)
'                        If InpTot(0) = "" Then
'                            MsgRg.Value = "1st element of InpLy cannot be empty"
'                            GoTo X
'                        End If                    '<== ShwMsg if no Input
''                        Ay = Array(C2_Tot_Sum, C2_Tot_Avg, C2_Tot_Cnt)
'                        For J = 0 To UB(InpTot)
'                            T = InpTot(J)
'                            If Not AyHas(Ay, T) Then
'                                MsgRg.Value = "*Tot column must be one of these [" & JnSpc(Ay) & "]"
'                                GoTo X
'                            End If
'Nxt:
'                        Next
'Dim InpFldLvs$():
'             InpFldLvs = VBarSy(InpFldLvsRg)
'                         If InpFldLvs(0) = "" Then
'                            MsgRg.Value = "1st element of InpFldLvs cannot be empty"
'                            GoTo X
'                         End If                    '<== ShwMsg if no Input
'Dim InpFny$():
'                InpFny = VBarSy(InpFnyRg)
'                         If InpFny(0) = "" Then
'                            MsgRg.Value = "1st element of InpFld cannot be empty"
'                            GoTo X
'                         End If                    '<== ShwMsg if no Input
'
'Dim DifSz As Boolean
'                 DifSz = Sz(InpTot) <> Sz(InpFldLvs)
'                         If DifSz Then
'                            MsgRg.Value = "FldLvs & *Tot are dif sz"
'                            GoTo X
'                         End If
'
'                         AyRgV IntAy_ByU(UB(InpFldLvs)), InpLxRg     '<== Put Lx: 0..
'
'Dim Fny$()
'Stop
''Run
''Dim Sum As LVF
''Dim Avg As LVF
''Dim Cnt As LVF
''                         For J = 0 To UB(InpFldLvs)
''                            T = InpTot(J)
''                            Select Case T
'''                            Case C2_Tot_Sum: Sum  LVF_Push3 Sum, J, C2_Tot_Sum, InpFldLvs(J))
'''                            Case C2_Tot_Avg: Avg  LVF_Push3 Avg, J, C2_Tot_Avg, InpFldLvs(J))
'''                            Case C2_Tot_Cnt: Cnt  LVF_Push3 Cnt, J, C2_Tot_Cnt, InpFldLvs(J))  '<== Calling
''                            Case Else:
''                            Stop
''                            End Select
''                         Next
''Put Rslt
'                         OupRg.Clear
'                         OupRg.Font.Name = "Courier New"
''Dim RsltLy$():  RsltLy = ZZ51_RsltLy(Sum, Avg, Cnt, Fny)
''                         AyRgH RsltLy, OupRg
'X:
'    IsInChg = False
End Sub
