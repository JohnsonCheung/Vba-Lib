VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Str$, Som As Boolean

Property Get ToStr$(Optional W% = 50)
If Som Then
    If Len(Str) < W Then
        ToStr = "*SomStr " & Str
    Else
        ToStr = "*SomStr " & AlignL(Str, 50)
    End If
Else
    ToStr = "*NoStr"
End If
End Property

Sub Dmp()
Debug.Print ToStr
End Sub

Friend Property Get Init(S) As SomStr
Str = S
Som = True
Set Init = Me
End Property

Function IsEq(A As SomStr) As Boolean
If Som <> A.Som Then Exit Function
If Str <> A.Str Then Exit Function
IsEq = True
End Function
