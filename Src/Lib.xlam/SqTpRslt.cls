VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqTpRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sql$, Er As New Er, Tp$
Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("SqTpRslt", "Sql Tp Er", StrObj(Sql), StrObj(Tp), Er)
End Property
