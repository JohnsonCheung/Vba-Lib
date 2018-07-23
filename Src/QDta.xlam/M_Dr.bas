Attribute VB_Name = "M_Dr"
Option Explicit

Function DrExpLinesCol(Dr, LinesColIx%) As Variant()
Dim B$()
    B = SplitCrLf(CStr(Dr(LinesColIx)))
Dim O()
    Dim IDr
        IDr = Dr
    Dim I
    For Each I In B
        IDr(LinesColIx) = I
        Push O, IDr
    Next
DrExpLinesCol = O
End Function

Function DrLin$(Dr, Wdt%())
Dim UDr%
   UDr = UB(Dr)
Dim O$()
   Dim U1%: U1 = UB(Wdt)
   ReDim O(U1)
   Dim W, V
   Dim J%, V1$
   J = 0
   For Each W In Wdt
       If UDr >= J Then V = Dr(J) Else V = ""
       V1 = AlignL(V, W)
       O(J) = V1
       J = J + 1
   Next
DrLin = Quote(Join(O, " | "), "| * |")
End Function
