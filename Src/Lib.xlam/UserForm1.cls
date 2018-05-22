VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Source Code"
   ClientHeight    =   4830
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7330
   Enabled         =   0   'False
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LV_AddAy(Sy$())
Dim I
For Each I In Sy
    Me.ListView1.ListItems.Add , I
Next
End Sub

Private Sub UserForm_Initialize()
MsgBox CurMdNm
LV_AddAy MdMthNy(CurMd)
End Sub

Private Sub UserForm_Resize()
With Me.ListView1
    .Top = 1
    .Left = 1
    .Height = Me.Height = 10
    .Width = Me.Width = 10
End With
End Sub
