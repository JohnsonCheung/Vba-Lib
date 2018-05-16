VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AddinInstall()
MsgBox "Install"
Z_EnsCustomUI
End Sub

Private Sub Workbook_Open()
MsgBox "Open"
Z_EnsCustomUI
End Sub

Private Sub Workbook_AddinUninstall()
On Error Resume Next
Kill ZFfn
MsgBox "Lib is unloaded"
End Sub

Private Sub Z_EnsCustomUI()
If Not ZIsExist Then ZCrtOne: MsgBox "Not exist and create": Exit Sub
If ZIsSam Then Exit Sub
ZCrtOne
MsgBox "Diff and created"
End Sub

Private Function ZIsExist() As Boolean
ZIsExist = FfnIsExist(ZFfn)
End Function

Private Sub ZCrtOne()
StrWrt ZCxt, ZFfn
End Sub
Private Function ZCxt$()
ZCxt = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui""><mso:ribbon><mso:qat/><mso:tabs><mso:tab id=""mso_c1.7DEF5EEE"" label=""New Tab"" insertBeforeQ=""mso:TabInsert""><mso:group id=""mso_c2.7DEF5EEE"" label=""New Group"" autoScale=""true""><mso:control idQ=""mso:FontSizeDecrease"" visible=""true""/></mso:group></mso:tab></mso:tabs></mso:ribbon></mso:customUI>"
End Function
Private Function ZFfn$()
ZFfn = "C:\Users\user\AppData\Local\Microsoft\Office\Excel.officeUI"
End Function
Private Function ZIsSam() As Boolean
ZIsSam = ZCxt = ZCxtFmFil
End Function
Private Function ZCxtFmFil$()
If ZIsExist Then
    ZCxtFmFil = FtLines(ZFfn)
End If
End Function

