VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AFlds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Flds As Fields

Property Get Dr() As Variant()
Dim I As Adodb.Field
Dim O()
For Each I In Flds
   Push O, I.Value
Next
Dr = O
End Property

Property Get Fny() As String()
Dim O$()
Dim F As Adodb.Field
For Each F In Flds
    Push O, F.Name
Next
Fny = O
End Property

Friend Property Get Init(A As Fields) As AFlds
Set Flds = A
Set Init = Me
End Property
