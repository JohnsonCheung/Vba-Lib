Attribute VB_Name = "M_WinOf"
Option Explicit
Function WinOf_BrwObj() As VBIDE.Window
Set WinOf_BrwObj = WinTy_Win(vbext_wt_Browser)
End Function
Function WinOf_Imm() As VBIDE.Window
Set WinOf_Imm = WinTy_Win(vbext_wt_Immediate)
End Function
Sub WinOf_Imm_Clr()
With WinOf_Imm
    .SetFocus
    .Visible = True
End With
Interaction.SendKeys "^{HOME}^+{END} ", True
End Sub
Sub WinOf_Imm_Cls()
DoEvents
With WinOf_Imm
    .Visible = False
End With
DoEvents
WinOf_BrwObj.SetFocus
'Interaction.SendKeys "^{F4}", True
'DoEvents
End Sub
