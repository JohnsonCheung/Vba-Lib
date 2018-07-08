VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoFmtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_BdrCnoAy%()
Private A_HidCnoAy%()
Private A_SumCny$(), A_AvgCny$(), A_CntCny$() 'Cny - Col Name Array
Private A_Bet() As CnoVal
Private A_Cor() As CnoVal
Private A_Fml() As CnoVal
Private A_Fmt() As CnoVal
Private A_Lbl() As CnoVal
Private A_Lvl() As CnoVal
Private A_Tit() As CnoVal
Private A_Wdt() As CnoVal

Property Get AvgCny() As String()
AvgCny = A_AvgCny
End Property

Property Get BdrCnoAy() As Integer()
BdrCnoAy = A_BdrCnoAy
End Property

Property Get CntCny() As String()
CntCny = A_CntCny
End Property

Property Get HidCnoAy() As Integer()
HidCnoAy = A_HidCnoAy
End Property

Property Get SumCny() As String()
SumCny = A_SumCny
End Property

Property Get ToStr$()
Stop '
'Dim O$()
'Push O, Tag("SumCny", SyObj(ASumCny))
'Push O, Tag("AvgCny", SyObj(A_AvgCny))
'Push O, Tag("CntCny", SyObj(A_CntCny))
'Push O, Tag("BdrCnoAy", IntAyObj(A_BdrCnoAy))
'Push O, Tag("HidCnoAy", IntAyObj(A_HidCnoAy))
'Push O, Tag("Bet", Bet)
'Push O, Tag("Cor", Cor)
'Push O, Tag("Fml", Fml)
'Push O, Tag("Fmt", Fmt)
'Push O, Tag("Lbl", Lbl)
'Push O, Tag("Lvl", Lvl)
'Push O, Tag("Tit", Tit)
'Push O, Tag("Wdt", Wdt)
'Const A$ = "SumCny AvgCny CntCny " & _
'"BdrCnoAy HidCnoAy " & _
'"Bet Cor Fml Fmt " & _
'"Lbl Lvl Tit Wdt"
'ToStr = Tag_NyStr_ObjAp("FmtWs", A, _
'SyObj(A_SumCny), SyObj(A_SumCny), SyObj(A_CntCny), _
'IntAyObj(A_BdrCnoAy), IntAyObj(A_HidCnoAy), _
'Bet, Cor, Fml, Fmt, _
'Lbl, Lvl, Tit, Wdt)
End Property

Friend Property Get Init(SumCny$(), AvgCny$(), _
CntCny$(), BdrCnoAy%(), HidCnoAy%(), _
Bet() As CnoVal, Cor() As CnoVal, _
Fml() As CnoVal, Fmt() As CnoVal, _
Lbl() As CnoVal, Lvl() As CnoVal, _
Tit() As CnoVal, Wdt() As CnoVal) As LoFmtr
A_SumCny = SumCny
A_AvgCny = AvgCny
A_CntCny = CntCny
A_BdrCnoAy = BdrCnoAy
A_HidCnoAy = HidCnoAy
A_Bet = Bet
A_Cor = Cor
A_Fml = Fml
A_Fmt = Fmt
A_Lbl = Lbl
A_Lvl = Lvl
A_Tit = Tit
A_Wdt = Wdt
Set Init = Me
End Property
