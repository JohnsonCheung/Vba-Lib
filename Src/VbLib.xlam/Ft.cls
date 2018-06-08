VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Ft$
Friend Property Get Init(A) As Ft
Ft = A
Set Init = Me
End Property

Sub Brw()
'Shell "code.cmd """ & Ft & """", vbHide
Shell "notepad.exe """ & Ft & """", vbMaximizedFocus
End Sub

Function FtLines$(Ft)
FtLines = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
End Function

Function FtDic(Ft) As Dictionary
Stop
'Set FtDic = Ly(FtLy(Ft)).Dic
End Function

Function FtLy(Ft) As String()
Dim F%: F = FtOpnInp(Ft)
Dim L$, O$()
While Not EOF(F)
    Line Input #F, L
    Push O, L
Wend
Close #F
FtLy = O
End Function

Function FtOpnApp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
FtOpnApp = O
End Function

Function FtOpnInp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FtOpnInp = O
End Function

Function FtOpnOup%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FtOpnOup = O
End Function

Sub PthBrw(P)
Shell "Explorer """ & P & """", vbMaximizedFocus
End Sub

