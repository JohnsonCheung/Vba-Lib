Attribute VB_Name = "M_Blk"
Option Explicit
Property Get LnxAy_BlkTyStr$(A() As Lnx)
Dim Ly$(): Ly = LnxAy_ZIs(Ly)
Dim O$
Select Case True
Case ZIsPm(Ly): O = "PM"
Case ZIsSw(Ly): O = "SW"
Case ZIsRm(Ly): O = "RM"
Case ZIsSq(Ly): O = "SQ"
Case Else: O = "ER"
End Select
ZBlkTyStr = O
End Property
Private Property Get ZIsSw(Ly$()) As Boolean
ZIsSw = ZIsHasMajPfx(Ly, "?")
End Property

Private Property Get ZIsPm(Ly$()) As Boolean
ZIsPm = ZIsHasMajPfx(Ly, "%")
End Property

Private Property Get ZIsRm(Ly$()) As Boolean
ZIsRm = AyIsEmp(Ly)
End Property

Property Get ZIsSq(Ly$()) As Boolean
If AyIsEmp(Ly) Then Exit Property
Dim L$: L = Ly(0)
Dim Sy$(): Sy = SslSy("?SEL SEL ?SELDIS SELDIS UPD DRP")
If HasOneOfPfxIgnCas(L, Sy) Then ZIsSq = True: Exit Property
End Property


