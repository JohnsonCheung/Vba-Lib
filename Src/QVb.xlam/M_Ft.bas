Attribute VB_Name = "M_Ft"
Option Explicit

Function FtDic(A) As Dictionary
Set FtDic = LyDic(FtLy(A))
End Function

Function FtLines$(A)
FtLines = Fso.GetFile(A).OpenAsTextStream.ReadAll
End Function

Function FtLy(A) As String()
FtLy = SplitLines(FtLines(A))
End Function

Function FtOpnApp%(A)
Dim O%: O = FreeFile(1)
Open A For Append As #O
FtOpnApp = O
End Function

Function FtOpnInp%(A)
Dim O%: O = FreeFile(1)
Open A For Input As #O
FtOpnInp = O
End Function

Function FtOpnOup%(A)
Dim O%: O = FreeFile(1)
Open A For Output As #O
FtOpnOup = O
End Function

Sub FtBrw(A)
Shell "code.cmd """ & A & """", vbHide
'Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub
