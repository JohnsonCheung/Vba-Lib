VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3VF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Nam As New LABCs
Public Hid As New LABCs
Public Fny As New LABCs
Public Bdr As New P3Bdr
Public Cor  As New LABCs
Public Fmt  As New LABCs
Public Lvl  As New LABCs
Public Tot  As New P3Tot
Public Wdt  As New LABCs

Private Sub Class_Initialize()
Nam.InitByT1 C1_Lo, IsVF:=True
Hid.InitByT1 C1_Lo, IsVF:=True
Fny.InitByT1 C1_Lo, IsVF:=True
Bdr.BdrL.InitByT1 C1_Bdr, IsVF:=True
Bdr.BdrC.InitByT1 C1_Bdr, IsVF:=True
Bdr.BdrR.InitByT1 C1_Bdr, IsVF:=True
Cor.InitByT1 C1_Cor, IsVF:=True
Fmt.InitByT1 C1_Fmt, IsVF:=True
Lvl.InitByT1 C1_Lvl, IsVF:=True
Tot.Cnt.InitByT1 C1_Tot, IsVF:=True
Tot.Avg.InitByT1 C1_Tot, IsVF:=True
Tot.Sum.InitByT1 C1_Tot, IsVF:=True
Wdt.InitByT1 C1_Wdt, IsVF:=True
End Sub
