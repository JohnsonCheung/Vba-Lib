VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TblNm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_TblNm$
Friend Property Get Init(TblNm) As TblNm
If Not IsNm(TblNm) Then PmEr
B_TblNm = TblNm
Set Init = Me
End Property

Property Get DrpStmts$()
DrpStmts = FmtQQ("Drop Table `?`", B_TblNm)
End Property
