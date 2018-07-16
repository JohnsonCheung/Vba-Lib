Attribute VB_Name = "M_Cur"
Option Explicit
Property Get CurWb() As Workbook
Set CurWb = Excel.Application.ActiveWorkbook
End Property


