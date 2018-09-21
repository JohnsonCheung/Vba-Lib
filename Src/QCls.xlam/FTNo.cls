VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fmno%, Tono% ' From-Number, To-Number.  Number is starting from 1
Friend Property Get Init(Fmno%, Tono%) As FTNo
If Fmno <= 0 Then Stop
If Tono <= 0 Then Stop
If Fmno - 1 > Tono Then Stop
Me.Fmno = Fmno
Me.Tono = Tono
Set Init = Me
End Property

Property Get ToStr$()
ToStr = "FTNo(Fm(" & Fmno & ") To(" & Tono & "))"
End Property
