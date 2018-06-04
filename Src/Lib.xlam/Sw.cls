VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public FldDic As New Dic, TblDic As New Dic

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("Sw", "FldDic TblDic", FldDic, TblDic)
End Property
