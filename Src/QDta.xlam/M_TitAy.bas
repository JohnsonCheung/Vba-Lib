Attribute VB_Name = "M_TitAy"
Option Explicit

Function TitAy_Sq(TitAy$()) As Variant()
Dim UFld%: UFld = UB(TitAy)
Dim ColVBar()
    ReDim ColVBar(UFld)
    Dim J%
    For J = 0 To UFld
        ColVBar(J) = AyTrim(SplitVBar(TitAy(J)))
    Next
Dim NRow%
    Dim M%, VBar$()
    For J = 0 To UB(ColVBar)
        VBar = ColVBar(J)
        M = Sz(VBar)
        If M > NRow Then NRow = M
    Next
Dim O()
    Dim I%
    ReDim O(1 To NRow, 1 To UFld + 1)
    For J = 0 To UFld
        VBar = ColVBar(J)
        For I = 0 To UB(VBar)
            O(I + 1, J + 1) = VBar(I)
        Next
    Next

End Function

Private Sub ZZ_TitAy_Sq()
Dim A$()
Push A, "ksdf | skdfj  |skldf jf"
Push A, "skldf|sdkfl|lskdf|slkdfj"
Push A, "askdfj|sldkf"
Push A, "fskldf"
SqBrw TitAy_Sq(A)
End Sub
