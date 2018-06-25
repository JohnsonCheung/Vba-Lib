Attribute VB_Name = "M_Ft"
Option Explicit

Property Get FtDic(A) As Dictionary
Set FtDic = LyDic(FtLy(A))
End Property

Property Get FtLines$(A)
FtLines = Fso.GetFile(A).OpenAsTextStream.ReadAll
End Property

Property Get FtLy(A) As String()
FtLy = SplitLines(FtLines(A))
End Property

Property Get FtOpnApp%(A)
Dim O%: O = FreeFile(1)
Open A For Append As #O
FtOpnApp = O
End Property

Property Get FtOpnInp%(A)
Dim O%: O = FreeFile(1)
Open A For Input As #O
FtOpnInp = O
End Property

Property Get FtOpnOup%(A)
Dim O%: O = FreeFile(1)
Open A For Output As #O
FtOpnOup = O
End Property

Sub FtBrw(A)
Shell "code.cmd """ & A & """", vbHide
'Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub
