Attribute VB_Name = "XlsXls"
Option Explicit
Function Xls() As Excel.Application
Static Y As Excel.Application
On Error GoTo X
Dim A$: A = Y.Name
Set Xls = Y
Exit Function
X:
Set Y = New Excel.Application
Set Xls = Y
End Function
Function XlsHasAddInFn(A As Excel.Application, AddInFn) As Boolean
Dim I As Excel.AddIn
Dim N$: N = UCase(AddInFn)
For Each I In A.AddIns
    If UCase(I.Name) = N Then XlsHasAddInFn = True: Exit Function
Next
End Function
Function XlsAddIn(A As Excel.Application, FxaNm) As Excel.AddIn
Dim I As Excel.AddIn
For Each I In A.AddIns
    If StrIsEq(I.Name, FxaNm & ".xlam") Then Set XlsAddIn = I
Next
End Function
Sub XlsAddFxaNm(A As Excel.Application, FxaNm$)
Dim F$: F = FxaNm_Fxa(FxaNm)
If F = "" Then Exit Sub
A.AddIns.Add FxaNm_Fxa(FxaNm)
End Sub
Sub XlsVis(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub
