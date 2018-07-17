Attribute VB_Name = "M_Obj"
Option Explicit

Property Get ObjNm$(O)
ObjNm = CallByName(O, "Name", VbGet)
End Property
Property Get ObjCompoundPrp$(Obj, PrpSsl$)
Dim Ny$(): Ny = SslSy(PrpSsl)
Dim O$(), I
For Each I In Ny
    Push O, CallByName(Obj, CStr(I), VbGet)
Next
ObjCompoundPrp = Join(O, "|")
End Property

Property Get ObjPrp(Obj, PrpPth$)
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim Ny$()
    Ny = Split(PrpPth, ".")
Dim O
    Dim J%, U%
    Set O = Obj
    U = UB(Ny)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, Ny(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next

ObjPrp = CallByName(O, Ny(U), VbGet) ' Last Prp may be non-object, so must use 'Asg'
End Property


Private Sub ZZZ_ObjCompoundPrp()
Dim Act$: Act = ObjCompoundPrp(Excel.Application.VBE.ActiveVBProject, "FileName Name")
Ass Act = "C:\Users\user\Desktop\Vba-Lib-1\QVb.xlam|QVb"
End Sub

