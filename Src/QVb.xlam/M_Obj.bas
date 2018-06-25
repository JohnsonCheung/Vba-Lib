Attribute VB_Name = "M_Obj"
Option Explicit

Property Get ObjNm$(O)
ObjNm = CallByName(O, "Name", VbGet)
End Property
