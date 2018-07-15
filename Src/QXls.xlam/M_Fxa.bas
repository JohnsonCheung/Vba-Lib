Attribute VB_Name = "M_Fxa"
Option Explicit

Property Get FxaCrt(A) As Excel.Application
If FfnIsExist(A) Then MsgBox FmtQQ("Fxa.Crt: Fxa(?) exist", A): Exit Property
Dim Wb As Workbook, X As Excel.Application
   Set X = New Excel.Application
   Set Wb = X.Workbooks.Add
X.VBE.VBProjects(1).Name = FfnFnn(A)
Wb.SaveAs A, XlFileFormat.xlOpenXMLAddIn
Wb.Close
X.AddIns.Add A
Set FxaCrt = X
End Property

Sub ZZ__Tst()
ZZ_FxaCrt
End Sub

Private Sub ZZ_FxaCrt()
Stop '
'Dim Act As Excel.Application
'A = TmpFxa
'Set Act = FxaCrt(A)
'Act.Visible = True
'Stop
End Sub
