Attribute VB_Name = "DtaDrs"
Option Explicit
Sub DrsBrw(A As Drs)
WsVis DrsWs(A)
End Sub
Function DrsInsCol(A As Drs, ColNm$, C) As Drs
Set DrsInsCol = Drs(AyIns(A.Fny, ColNm), DryInsCol(A.Dry, C))
End Function
Function DrsWhColEq(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, C)
Set DrsWhColEq = Drs(Fny, DryWhColEq(A.Dry, Ix, V))
End Function
Function DrsWhColGt(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, C)
Set DrsWhColGt = Drs(Fny, DryWhColGt(A.Dry, Ix, V))
End Function
Function DrsAddValIdCol(A As Drs, ColNm$) As Drs
Dim Fny$(), ValCol%
Fny = A.Fny
PushI Fny, ColNm & "Id"
PushI Fny, ColNm & "Cnt"
Set DrsAddValIdCol = Drs(Fny, DryAddValIdCol(A.Dry, AyIx(Fny, ColNm)))
End Function
Function DrsInsColBef(A As Drs, C$, FldNm$) As Drs
Set DrsInsColBef = DrsInsColXxx(A, C, FldNm, False)
End Function
Function DrsInsColAft(A As Drs, C$, FldNm$) As Drs
Set DrsInsColAft = DrsInsColXxx(A, C, FldNm, True)
End Function
Private Function DrsInsColXxx(A As Drs, C$, FldNm$, IsAft As Boolean) As Drs
Dim Fny$(), Dry(), Ix&, Fny1$()
Fny = A.Fny
Ix = AyIx(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Fny1 = AyIns(Fny, FldNm, CLng(Ix))
Dry = DryInsCol(A.Dry, Ix)
Set DrsInsColXxx = Drs(Fny1, Dry)
End Function
Function DrsKeyCntDic(A As Drs, K$) As Dictionary
Dim Dry(), O As New Dictionary, Fny$(), Dr, Ix%, KK$
Fny = A.Fny
Ix = AyIx(Fny, K)
Dry = A.Dry
If Sz(Dry) > 0 Then
    For Each Dr In A.Dry
        KK = Dr(Ix)
        If O.Exists(KK) Then
            O(KK) = O(KK) + 1
        Else
            O.Add KK, 1
        End If
    Next
End If
Set DrsKeyCntDic = O
End Function
Function DrsWs(A As Drs) As Worksheet
Dim O As Worksheet
Set O = NewWs
AyRgH A.Fny, WsA1(O)
RgLo RgIncTopR(SqRg(DrySq(A.Dry), WsRC(O, 2, 1)))
Set DrsWs = O
End Function
Function DrsSq(A As Drs) As Variant()
Dim NCol&, NRow&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    NCol = Max(DryNCol(Dry), Sz(Fny))
    NRow = Sz(Dry)
Dim O()
ReDim O(1 To 1 + NRow, 1 To NCol)
Dim C&, R&, Dr()
    For C = 1 To Sz(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NRow
        Dr = A(R - 1)
        For C = 1 To Min(Sz(Dr), NCol)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
DrsSq = O
End Function
Function DrsGpFlat(A As Drs, K$, G$) As Drs
Dim Fny0$, Dry(), S$, N%, Ix%()
Ix = AyIxAyI(A.Fny, Array(K, G))
Dry = DryGpFlat(A.Dry, Ix(0), Ix(1))
N = DryNCol(Dry) - 2
S = LblSeqSsl(G, N)
Fny0 = FmtQQ("? N ?", K, S)
Set DrsGpFlat = Drs(Fny0, Dry)
End Function
