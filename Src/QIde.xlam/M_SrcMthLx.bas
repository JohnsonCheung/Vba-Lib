Attribute VB_Name = "M_SrcMthLx"
Option Explicit
Function SrcMthLx_BdyLy(A$(), MthLx) As String()
Dim ToLx%: ToLx = SrcMthLx_ToLx(A, MthLx)
Dim FmLx%: FmLx = SrcMthLx_MthRmkLx(A, MthLx)
Dim FT As FmTo
With FT
   .FmIx = FmLx
   .ToIx = ToLx
End With
Dim O$()
   O = AyWhFmTo(A, FT)
SrcMthLx_BdyLy = O
If AyLasEle(O) = "" Then Stop
End Function
Function SrcMthLx_MthRmkLx&(A$(), MthLx)
Dim M1&
    Dim J&
    For J = MthLx - 1 To 0 Step -1
        If SrcLin_IsCd(A(J)) Then
            M1 = J
            GoTo M1IsFnd
        End If
    Next
    M1 = -1
M1IsFnd:
Dim M2&
    For J = M1 + 1 To MthLx - 1
        If Trim(A(J)) <> "" Then
            M2 = J
            GoTo M2IsFnd
        End If
    Next
    M2 = MthLx
M2IsFnd:
SrcMthLx_MthRmkLx = M2
End Function
Private Sub ZZ_SrcMthLx_MthRmkLx()
Dim ODry()
    Dim Src$(): Src = MdSrc(Md("IdeSrcLin"))
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, Lin
    For Each Lin In Src
        IsMth = ""
        RmkLx = ""
        If SrcLin_IsMth(Lin) Then
            If Lx = 482 Then Stop
            IsMth = "*Mth"
            RmkLx = SrcMthLx_MthRmkLx(Src, Lx)

        End If
        Dr = Array(IsMth, RmkLx, Lin)
        Push ODry, Dr
        Lx = Lx + 1
    Next
DrsBrw Drs("Mth RmkLx Lin", ODry)
End Sub
