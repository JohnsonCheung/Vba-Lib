Attribute VB_Name = "VbObj"
Option Explicit
Function ObjPrp(Obj, PrpNm)
On Error Resume Next
Asg CallByName(Obj, PrpNm, VbGet), ObjPrp
End Function
Function ObjToStr$(A)
If Not IsObject(A) Then Stop
On Error GoTo X
ObjToStr = A.ToStr: Exit Function
X: ObjToStr = QuoteSqBkt(TypeName(A))
End Function
