VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoPrivateClassCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Friend Property Get Cn(A As ADODB.Connection) As AdoCn
Dim O As New AdoCn
Set O.Cn = A
Set Cn = O
End Property

Friend Property Get Flds(A As ADODB.Fields) As AdoFlds
Dim O As New AdoFlds
Set O.Flds = A
Set Flds = O
End Property

Friend Property Get Rs(A As ADODB.Recordset) As AdoRs
Dim O As New AdoRs
Set O.Rs = A
Set Rs = O
End Property
