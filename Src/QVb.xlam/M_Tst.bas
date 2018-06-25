Attribute VB_Name = "M_Tst"
Option Explicit
Type ThowMsgOrStr
    Som As Boolean
    Str As String
    ThowMsg As String
End Type
Type ThowMsgOrSy
    Som As Boolean
    Sy() As String
    ThowMsg As String
End Type
Type ThowMsgOrInt
    Som As Boolean
    Int As Integer
    ThowMsg As String
End Type
Type ThowMsgOrVar
    Som As Boolean
    V As Variant
    ThowMsg As String
End Type

Property Get TMOIntSomInt(I%) As ThowMsgOrInt
TMOIntSomInt.Som = True
TMOIntSomInt.Int = I
End Property

Property Get TMOIntThowMsg(ThowMsg$) As ThowMsgOrInt
TMOIntThowMsg.ThowMsg = ThowMsg
End Property

Property Get TMOStrDmp(A As ThowMsgOrStr, Optional Nm$ = "ThowMsgOrStr")
With A
    Debug.Print Nm$; " = ";
    Debug.Print IIf(.Som, "SomStr ", "SomThowMsg ");
    Debug.Print IIf(.Som, .Str, .ThowMsg)
End With
End Property

Property Get TMOStrSomStr(Str$) As ThowMsgOrStr
TMOStrSomStr.Som = True
TMOStrSomStr.Str = Str
End Property

Property Get TMOStrThowMsg(ThowMsg$) As ThowMsgOrStr
TMOStrThowMsg.ThowMsg = ThowMsg
End Property

Property Get TMOSySomSy(Sy$()) As ThowMsgOrSy
TMOSySomSy.Som = True
TMOSySomSy.Sy = Sy
End Property

Property Get TMOSySomThowMsg(ThowMsg$) As ThowMsgOrSy
TMOSySomThowMsg.ThowMsg = ThowMsg
End Property

Property Get TstResFdr$(Fdr$)
Dim O$
    O = TstResPth & Fdr & "\"
    PthBrw O
TstResFdr = O
End Property

Property Get TstResPth$()
Dim O$
    Stop '
'    O = CurPjx.SrcPth & "TstRes\"
    PthEns O
TstResPth = O
End Property

Sub TstResFdrBrw(Fdr$)
'PthBrw TstResFdr(Fdr)).Brw
Stop '
End Sub

Sub TstResPthBrw()
PthBrw TstResPth
End Sub
