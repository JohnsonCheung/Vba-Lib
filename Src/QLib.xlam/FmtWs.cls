VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmtWs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Bdr%() ' CnoAy
Private B_Hid%() ' CnoAy
Private B_SumCny$(), B_AvgCny$(), B_CntCny$() 'Cny - Col Name Array
Private B_Bet() As CnoVal
Private B_Cor() As CnoVal
Private B_Fml() As CnoVal
Private B_Fmt() As CnoVal
Private B_Lbl() As CnoVal
Private B_Tit() As CnoVal
Private B_Lvl() As CnoVal
Private B_Wdt() As CnoVal
Property Get Init() As FmtWs

End Property
Property Get Lvl() As CnoVal()
Lvl = B_Lvl
End Property
Property Get Hid() As Integer()
Hid = B_Hid
End Property
Property Get Wdt() As CnoVal()
Wdt = B_Wdt
End Property
Property Get Bdr() As Integer()
Bdr = B_Bdr
End Property
Property Get Fmt() As CnoVal()
Fmt = B_Fmt
End Property
Property Get Fml() As CnoVal()
Fml = B_Fml
End Property
Property Get Lbl() As CnoVal()
Lbl = B_Lbl
End Property
Property Get Cor() As CnoVal()

End Property
Property Get ToStr$()
Dim O$()
Stop '
'ToStr = Tag_NyStr_ObjAp("FmtWs", "Lvl Hid Wdt Bdr Cor Fmt   Lbl Tit    Bet Fmt Sum Avg Cnt", _
'    Lvl, Hid, Wdt, Bdr, Cor, Fmt, _
'    Lbl, Tit, _
'    Bet, Fml, Sum, Avg, Cnt)
End Property

