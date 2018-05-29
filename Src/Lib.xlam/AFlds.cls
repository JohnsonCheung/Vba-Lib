VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AFlds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Flds As AdoDb.Fields
Function Dr() As Variant()
Dim I As AdoDb.Field
Dim O()
For Each I In Flds
   Push O, I.Value
Next
Dr = O
End Function

Function Fny() As String()
Dim O$()
Dim F As AdoDb.Field
For Each F In Flds
    Push O, F.Name
Next
Fny = O
End Function
