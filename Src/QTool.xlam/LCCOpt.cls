VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LCCOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Som As Boolean, LCC As LCC
Friend Property Get Init(LCC As LCC) As LCCOpt
Set Me.LCC = LCC
Som = True
Set Init = Me
End Property
