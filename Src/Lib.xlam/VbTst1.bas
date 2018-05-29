Attribute VB_Name = "VbTst1"
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

Function TMOIntSomInt(I%) As ThowMsgOrInt
TMOIntSomInt.Som = True
TMOIntSomInt.Int = I
End Function

Function TMOIntThowMsg(ThowMsg$) As ThowMsgOrInt
TMOIntThowMsg.ThowMsg = ThowMsg
End Function

Function TMOStrDmp(A As ThowMsgOrStr, Optional Nm$ = "ThowMsgOrStr")
With A
    Debug.Print Nm$; " = ";
    Debug.Print IIf(.Som, "SomStr ", "SomThowMsg ");
    Debug.Print IIf(.Som, .Str, .ThowMsg)
End With
End Function

Function TMOStrSomStr(Str$) As ThowMsgOrStr
TMOStrSomStr.Som = True
TMOStrSomStr.Str = Str
End Function

Function TMOStrThowMsg(ThowMsg$) As ThowMsgOrStr
TMOStrThowMsg.ThowMsg = ThowMsg
End Function

Function TMOSySomSy(Sy$()) As ThowMsgOrSy
TMOSySomSy.Som = True
TMOSySomSy.Sy = Sy
End Function

Function TMOSySomThowMsg(ThowMsg$) As ThowMsgOrSy
TMOSySomThowMsg.ThowMsg = ThowMsg
End Function

Function TstResPth$()
Dim O$
    O = CurPj.SrcPth & "TstRes\"
    PthEns O
TstResPth = O
End Function
Sub TstResPthBrw()
PthBrw TstResPth
End Sub

Function TstResFdr$(Fdr$)
Dim O$
    O = TstResPth & Fdr & "\"
    PthEns O
TstResFdr = O
End Function

Sub TstResFdrBrw(Fdr$)
PthBrw TstResFdr(Fdr)
End Sub



Sub ZGenTstXXX()
Dim Qvbl$
Dim Lvs$
Qvbl = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
'Lvs = JnSpc(PjClsNy(CurPj))
Debug.Print Seed(Qvbl).Expand(Lvs)
End Sub
