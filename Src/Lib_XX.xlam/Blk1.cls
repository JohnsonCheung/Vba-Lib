VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blk1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Public Lnxs As Lnx1s
'Private B_BlkTyStr$
'Property Get BlkTyStr$()
'BlkTyStr = B_BlkTyStr
'End Property
'Friend Property Get Init(A As Lnx1s) As Blk1
'Set Lnxs = A
'B_BlkTyStr = ZBlkTyStr
'Set Init = Me
'End Property
'Private Function ZBlkTyStr$()
'Dim Ly$(): Ly = Lnxs.Ly
'Dim O$
'Select Case True
'Case LyIsPm(Ly): O = "PM"
'Case LyIsSw(Ly): O = "SW"
'Case LyIsRm(Ly): O = "RM"
'Case LyIsSq(Ly): O = "SQ"
'Case Else: O = "ER"
'End Select
'ZBlkTyStr = O
'End Function
'Private Function LyIsSw(A$()) As Boolean
'LyIsSw = LyHasMajPfx(A, "?")
'End Function
'
'Private Function LyIsPm(A$()) As Boolean
'LyIsPm = LyHasMajPfx(A, "%")
'End Function
'
'Private Function LyIsRm(A$()) As Boolean
'LyIsRm = AyIsEmp(A)
'End Function
'
'Private Function LyIsSq(A$()) As Boolean
'If AyIsEmp(A) Then Exit Function
'Dim L$: L = A(0)
'Dim Sy$(): Sy = LvsSy("?SEL SEL ?SELDIS SELDIS UPD DRP")
'If HasOneOfPfxIgnCas(L, Sy) Then LyIsSq = True: Exit Function
'End Function
'
