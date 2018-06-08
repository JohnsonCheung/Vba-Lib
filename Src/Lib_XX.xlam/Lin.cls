VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Public A$
'Friend Property Get Init(Lin) As Lin
'A = Lin
'Set Init = Me
'End Property
'Private Sub ZZ_RmvT1()
'Ass Vb.Lin("  df dfdf  ").RmvT1 = "dfdf"
'End Sub
'
'Property Get PfxErMsg$(Pfx$)
'If HasPfx(Lin, Pfx) Then Exit Property
'PfxErMsg = FmtQQ("First Char must be [?]", Pfx)
'End Property
'
'Property Get RmvT1$()
'RmvT1 = Brk1(Trim, " ").S2
'End Property
'Property Get Lin$()
'Lin = A
'End Property
'Property Get ShiftTerm$()
'With Brk1(A, " ")
'    ShiftTerm = .S1
'    A = .S2
'End With
'End Property
'
'Property Get T1$()
'T1 = Brk1(Trim, " ").S1
'End Property
'
'Property Get T2$()
'Dim L$: L = RmvT1
'T2 = Vb.Lin(L).T1
'End Property
'Sub AsgTRst(OTerm, ORst)
'With Vb.Lin(A)
'    OTerm = .ShiftTerm
'    ORst = .Lin
'End With
'End Sub
'
'Sub AsgTTRst(OTerm1, OTerm2, ORst)
'With Vb.Lin(A)
'    OTerm1 = .ShiftTerm
'    OTerm2 = .ShiftTerm
'    ORst = .Lin
'End With
'End Sub
'
'
'Property Get IsSngTerm() As Boolean
'With Brk1(Lin, " ")
'    IsSngTerm = .S1 <> "" And .S2 = ""
'End With
'End Property
'
'Property Get T1Rst() As T1Rst
'Dim O As T1Rst
'With Brk1(A, " ")
'    O.T1 = .S1
'    O.Rst = .S2
'End With
'T1Rst = O
'End Property
'
'Property Get LTrim$()
'LTrim = VBA.LTrim(Lin)
'End Property
'
'Property Get Trim$()
'Trim = VBA.Trim(Lin)
'End Property
'Property Get RmvDDRmk() As Lin
'Dim S$
'If HasDDRmk Then
'    S = ""
'Else
'    S = Lin
'End If
'Set RmvDDRmk = Vb.Lin(S)
'End Property
'Property Get HasDDRmk() As Boolean
'HasDDRmk = HasPfx(LTrim, "--")
'End Property
'
'Property Get IsEmp() As Boolean
'IsEmp = Trim = ""
'End Property
'
'Property Get Nm$()
'Dim J%
'If IsLetter(FstChr(Lin)) Then
'   For J = 2 To Len(Lin)
'       If Not IsNmChr(Mid(Lin, J, 1)) Then Exit For
'   Next
'   Nm = Left(Lin, J - 1)
'End If
'End Property
'
