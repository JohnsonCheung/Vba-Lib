VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Str$, Som As Boolean

Function IsEq(A As StrOpt) As Boolean
If Som <> A.Som Then Exit Function
If Str <> A.Str Then Exit Function
IsEq = True
End Function

Sub Dmp()
Debug.Print ToStr
End Sub

Friend Property Get Init(S) As StrOpt
Str = S
Som = True
Set Init = Me
End Property

Property Get ToStr$(Optional W% = 50)
If Som Then
    If Len(Str) < W Then
        ToStr = "*StrOpt " & Str
    Else
        ToStr = "*StrOpt " & AlignL(Str, 50)
    End If
Else
    ToStr = "*NoStr"
End If
End Property


