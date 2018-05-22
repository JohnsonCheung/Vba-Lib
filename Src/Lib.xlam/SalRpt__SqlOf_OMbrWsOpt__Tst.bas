Attribute VB_Name = "SalRpt__SqlOf_OMbrWsOpt__Tst"
Option Explicit
Private Type ZZ2_TstDta
    Exp As String
    BrkMbr As Boolean
    InclAdr As Boolean
    InclEmail As Boolean
    InclNm As Boolean
    InclPhone As Boolean
End Type

Private Sub ZZ2_Push(O() As ZZ2_TstDta, I As ZZ2_TstDta)

End Sub

Private Function ZZ2_TstDta0() As ZZ2_TstDta
With ZZ2_TstDta0
    .BrkMbr = False
    .Exp = ""
End With
End Function

Private Function ZZ2_TstDta10() As ZZ2_TstDta
With ZZ2_TstDta10
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta11() As ZZ2_TstDta
With ZZ2_TstDta11
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta12() As ZZ2_TstDta
With ZZ2_TstDta12
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta1() As ZZ2_TstDta
With ZZ2_TstDta1
    .BrkMbr = True
    .Exp = "Select|    JCMCode                                                        Mbr ,|    DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())      Age ,|    JCMSex                                                         Sex ,|    JCMStatus                                                      Sts ,|    JCMDist                                                        Dist,|    JCMArea                                                        Area|  Into #MbrDta|  From JCMember|  Where JCMDCode in (Select Mbr From #TxMbr)"
End With
End Function

Private Function ZZ2_TstDta2() As ZZ2_TstDta
With ZZ2_TstDta2
    .BrkMbr = True
    .InclAdr = True
    .Exp = "Select|    JCMCode                                                        Mbr ,|    DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())      Age ,|    JCMSex                                                         Sex ,|    JCMStatus                                                      Sts ,|    JCMDist                                                        Dist,|    JCMArea                                                        Area,|    Adr-Express-L1|      Adr-Expression-L2                                              Adr |  Into #MbrDta|  From JCMember|  Where JCMDCode in (Select Mbr From #TxMbr)"
End With
End Function

Private Function ZZ2_TstDta3() As ZZ2_TstDta
With ZZ2_TstDta3
    .BrkMbr = True
    .InclAdr = True
    .InclEmail = True
    .Exp = "Select|    JCMCode                                                        Mbr  ,|    DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())      Age  ,|    JCMSex                                                         Sex  ,|    JCMStatus                                                      Sts  ,|    JCMDist                                                        Dist ,|    JCMArea                                                        Area ,|    Adr-Express-L1|      Adr-Expression-L2                                              Adr  ,|    JCMEmail                                                       Email|  Into #MbrDta|  From JCMember|  Where JCMDCode in (Select Mbr From #TxMbr)"
End With
End Function

Private Function ZZ2_TstDta4() As ZZ2_TstDta
With ZZ2_TstDta4
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta5() As ZZ2_TstDta
With ZZ2_TstDta5
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta6() As ZZ2_TstDta
With ZZ2_TstDta6
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta7() As ZZ2_TstDta
With ZZ2_TstDta7
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta8() As ZZ2_TstDta
With ZZ2_TstDta8
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDta9() As ZZ2_TstDta
With ZZ2_TstDta9
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function ZZ2_TstDtaAy() As ZZ2_TstDta()
Dim O()  As ZZ2_TstDta
ZZ2_Push O, ZZ2_TstDta0
ZZ2_Push O, ZZ2_TstDta1
ZZ2_Push O, ZZ2_TstDta2
ZZ2_Push O, ZZ2_TstDta3
ZZ2_Push O, ZZ2_TstDta4
ZZ2_Push O, ZZ2_TstDta5
ZZ2_Push O, ZZ2_TstDta6
ZZ2_Push O, ZZ2_TstDta7
ZZ2_Push O, ZZ2_TstDta8
ZZ2_Push O, ZZ2_TstDta9
ZZ2_Push O, ZZ2_TstDta10
ZZ2_Push O, ZZ2_TstDta11
ZZ2_Push O, ZZ2_TstDta12
ZZ2_TstDtaAy = O
End Function

Private Sub ZZ2_TstDtaDmp0()
ZZ2_TstDtaDmp 0
End Sub

Private Sub ZZ2_TstDtaDmp1()
ZZ2_TstDtaDmp 1
End Sub

Private Sub ZZ2_TstDtaDmp2()
ZZ2_TstDtaDmp 2
End Sub

Private Sub ZZ2_TstDtaDmp3()
ZZ2_TstDtaDmp 3
End Sub

Private Sub ZZ2_TstDtaDmp(CasNo%)
'Dim D As New Dictionary
'Dim M  As ZZ2_Tstdta
'    Dim Ay()  As ZZ2_Tstdta
'    Ay = ZZ2_TstDtaAy
'    M = Ay(CasNo)
'With M
'    D.Add "BrkMbr", .BrkMbr
'    D.Add "*CaseNo", CasNo
'    D.Add "InclAdr", .InclAdr
'    D.Add "InclEmail", .InclEmail
'    D.Add "InclNm", .InclNm
'    D.Add "InclPhone", .InclPhone
'End With
'Dim Exp$
'Dim Act$
'    Exp = M.Exp
'    Act = SqLoFmtr_OMbrWsOpt(
'D.Add "**", IIf(Act = Exp, "Pass", "Fail")
'DicDmp DicSrt(D)
'If Act = Exp Then
'    Debug.Print "Act = Exp ======================================"
'    Debug.Print RplVbar(Act)
'Else
'    Debug.Print "Exp ========================================="
'    Debug.Print RplVbar(Exp)
'    Debug.Print "Act ========================================="
'    Debug.Print RplVbar(Act)
'End If
'Ass IsEq( Act, Exp
End Sub

Private Sub ZZ2_Tstr(A As ZZ2_TstDta)
With A
    Ass IsEq(SqLoFmtr_OMbrWsOpt(.BrkMbr, .InclNm, .InclAdr, .InclEmail, .InclPhone), .Exp)
End With
End Sub

Private Sub ZZ2__SqLoFmtr_MbrWsOpt__Tst()
Dim Ay()  As ZZ2_TstDta
    Ay = ZZ2_TstDtaAy
Dim J%
For J = 0 To 12
    ZZ2_Tstr Ay(J)
Next
End Sub
