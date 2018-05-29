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
Friend Property Get Rs(A As AdoDb.Recordset) As ARs
Dim O As New ARs
Set O.Rs = A
Set Rs = O
End Property
Friend Property Get Cn(A As AdoDb.Connection) As ACn
Dim O As New ACn
Set O.Cn = A
Set Cn = O
End Property
Friend Property Get Flds(A As AdoDb.Fields) As AFlds
Dim O As New AFlds
Set O.Flds = A
Set Flds = O
End Property

