Attribute VB_Name = "IdeMdSrt"
'Option Explicit
'Type MdSrtRpt
'MdIdxDt As Dt
'RptDic As Dictionary ' K is Module Name, V is DicCmpRsltLy
'End Type
'
'Sub MdSrt(A As CodeModule)
'Dim Nm$: Nm = MdNm(A)
'Debug.Print "Sorting: "; AlignL(Nm, 30); " ";
'Dim Ay(): Ay = Array("IdeMdSrt")
''Skip some md
'    If AyHas(Ay, Nm) Then
'        Debug.Print "<<<< Skipped"
'        Exit Sub
'    End If
'Dim NewLines$: NewLines = MdSrtedLines(A)
'Dim Old$: Old = MdLines(A)
''Exit if same
'    If Old = NewLines Then
'        Debug.Print "<== Same"
'        Exit Sub
'    End If
'Debug.Print "<-- Sorted";
''Delete
'    Debug.Print FmtQQ("<--- Deleted (?) lines", A.CountOfLines);
'    MdClr A, IsSilent:=True
''Add sorted lines
'    A.AddFromString NewLines
'    MdRmvEndBlankLines A
'    Debug.Print "<----Sorted Lines added...."
'End Sub
'
'Function MdSrtCmpLy(A As CodeModule) As String()
'MdSrtCmpLy = SrcSrtCmpLy(MdSrc(A))
'End Function
'
'Function MdSrtCmpRslt(A As CodeModule) As DCRslt
'MdSrtCmpRslt = SrcSrtCmpRslt(MdSrc(A))
'End Function
'
'Function MdSrtedLines$(A As CodeModule)
'MdSrtedLines = SrcSrtedLines(MdSrc(A))
'End Function
'
'Function PjMdSrtRpt(A As VBProject) As MdSrtRpt
''SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
'Dim MdAy() As CodeModule: 'MdAy = PjMdAy(A)
'Dim MdNy$(): MdNy = Oy(MdAy).Ny
'Dim LyAy()
'Dim IsSam$(), IsDif$(), Sam As Boolean
'    Dim J%, R As DCRslt
'    For J = 0 To UB(MdAy)
'        R = MdSrtCmpRslt(MdAy(J))
'        Push LyAy, DCRsltLy(R)
'        Sam = DCRsltIsSam(R)
'        Push IsSam, IIf(Sam, "*Sam", "")
'        Push IsDif, IIf(Sam, "", "*Dif")
'    Next
'With PjMdSrtRpt
'    Set .RptDic = AyPair_Dic(MdNy, LyAy)
'    .MdIdxDt = NewDt("Md-Bef-Aft-Srt", "Md Sam Dif", AyZipAp(MdNy, IsSam, IsDif))
'End With
'End Function
'
'Function PjSrtCmpRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
'Dim A1 As MdSrtRpt
'A1 = PjMdSrtRpt(A)
'Dim O As Workbook: Set O = Dix(A1.RptDic).Wb
'Dim Ws As Worksheet
'Set Ws = WbAddWs(O, "Md Idx", IsBeg:=True)
'Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
'LoCol_LnkWs Lo, "Md"
'If Vis Then WbVis O
'Set PjSrtCmpRptWb = O
'End Function
'
'Function SrcSrtCmpLy(A$()) As String()
'Dim R As DCRslt: R = SrcSrtCmpRslt(A)
'SrcSrtCmpLy = DCRsltLy(R)
'End Function
'
'Function SrcSrtCmpRslt(A$()) As DCRslt
'Dim B$(): B = SrcSrtedLy(A)
'Dim A1 As Dictionary
'Dim B1 As Dictionary
'Set A1 = SrcDic(A)
'Set B1 = SrcDic(B)
'SrcSrtCmpRslt = Dix(A1).Cmp(B1)
'End Function
'
'Function SrcSrtedBdyLines$(A$())
'If AyIsEmp(A) Then Exit Function
'Dim Drs As Drs
'   Drs = SrcMthDrs(A, WithBdyLines:=True)
'Dim MthLinesAy$()
'   MthLinesAy = DrsStrCol(Drs, "BdyLines")
'Dim I&()
'   Dim Ky$(): Ky = MthDrs_SortingKy(Drs)
'   I = AySrtInToIxAy(Ky)
'Dim O$()
'Dim J%
'   For J = 0 To UB(I)
'       Push O, vbCrLf & MthLinesAy(I(J))
'   Next
'SrcSrtedBdyLines = JnCrLf(O)
'End Function
'
'Function SrcSrtedLines$(A$())
'Dim O$(), A1$, A2$
'A1 = LinesEndTrim(SrcDclLines(A))
'A2 = SrcSrtedBdyLines(A)
'If LasChr(A2) = vbCr Then Stop
'PushNonEmp O, A1
'PushNonEmp O, A2
'SrcSrtedLines = Join(O, vbCrLf)
'End Function
'
'Function SrcSrtedLy(A$()) As String()
'SrcSrtedLy = SplitCrLf(SrcSrtedLines(A))
'End Function
'
'Private Sub ZZ_Dcl_BefAndAft_Srt()
'Const MdNm$ = "VbStrRe"
'Dim A$() ' Src
'Dim B$() ' Src->Srt
'Dim A1$() 'Src->Dcl
'Dim B1$() 'Src->Src->Dcl
''A = MdSrc(Md(MdNm))
'B = SrcSrtedLy(A)
'A1 = SrcDclLy(A)
'B1 = SrcDclLy(B)
'Stop
'End Sub
'
'Private Sub ZZ_PjSrtCmpRptWb()
'Dim O As Workbook: Set O = PjSrtCmpRptWb(CurPjx, Vis:=True)
'Stop
'End Sub
'
'Private Sub ZZ_SrcSrtCmpLy()
'AyBrw SrcSrtCmpLy(ZZSrc)
'End Sub
'
'Private Sub ZZ_SrcSrted()
''Dim Src$(): Src = MdSrc(Md("ThisWorkbook"))
''Dim Src1$(): Src1 = SrcSrtedLy(Src)
'Stop
'End Sub
'
'Private Function ZZSrc() As String()
''ZZSrc = MdSrc(Md("IdeMdSrt"))
'End Function
'
'Private Sub ZZ_MdSrt()
'Dim Md As CodeModule
'GoSub X0
'Exit Sub
'X0:
'    Dim I
''    For Each I In PjMdAy(CurPjx)
'        Set Md = I
'        If MdNm(Md) = "Str_" Then
'            GoSub Ass
'        End If
''    Next
'    Return
'X1:
'
'    Return
'Ass:
'    Debug.Print MdNm(Md); vbTab;
'    Dim BefSrt$(), AftSrt$()
'    BefSrt = MdLy(Md)
'    AftSrt = SplitCrLf(MdSrtedLines(Md))
'    If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then
'        Debug.Print "Is Same of before and after sorting ......"
'        Return
'    End If
'    If Not AyIsEmp(AftSrt) Then
'        If AyLasEle(AftSrt) = "" Then
'            Dim Pfx
'            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & MdNm(Md) & "=====")
'            AyBrw AyAddAp(Pfx, AftSrt)
'            Stop
'        End If
'    End If
'    Dim A$(), B$(), II
'    A = AyMinus(BefSrt, AftSrt)
'    B = AyMinus(AftSrt, BefSrt)
'    Debug.Print
'    If Sz(A) = 0 And Sz(B) = 0 Then Return
'    If Not AyIsEmp(AyRmvEmp(A)) Then
'        Debug.Print "Sz(A)=" & Sz(A)
'        AyBrw A
'        Stop
'    End If
'    If Not AyIsEmp(AyRmvEmp(B)) Then
'        Debug.Print "Sz(B)=" & Sz(B)
'        AyBrw B
'        Stop
'    End If
'    Return
'End Sub
'
'Private Sub ZZ_MdSrtedLines()
''StrBrw MdSrtedLines(Md("Md_"))
'End Sub
'
'Private Sub ZZ_SrcSrtedBdyLines()
'StrBrw SrcSrtedBdyLines(ZZSrc)
'End Sub
'
'Private Sub ZZ_SrcSrtedLines()
'StrBrw SrcSrtedLines(ZZSrc)
'End Sub
