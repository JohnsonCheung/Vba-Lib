VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Sy1$(), A_Sy2$()

Property Get Init(Sy1$(), Sy2$()) As SyPair
A_Sy1 = Sy1
A_Sy2 = Sy2
Set Init = Me
End Property

Property Get Sy1() As String()
Sy1 = A_Sy1
End Property

Property Get Sy2() As String()
Sy2 = A_Sy2
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("SyPair", "Sy1 Sy2", SyObj(Sy1), SyObj(Sy2))
End Property

Friend Sub ZZ__Tst()
ZZ_ToStr
End Sub

Private Sub ZZ_ToStr()
Dim A$(): A = VblLy("a df|df bdf|df dfdklf")
Dim Act$: Act = Ly_T1Rst_SyPair(A).ToStr
End Sub
