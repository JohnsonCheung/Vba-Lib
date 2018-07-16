Attribute VB_Name = "M_Max"
Option Explicit

Property Get MaxCol%()
Static C%
If C = 0 Then
    Dim Ws As Worksheet
    Set Ws = ActiveSheet
    Dim Cls As Boolean
    If IsNothing(Ws) Then
        Set Ws = NewWs
        Cls = True
    End If
    C = Ws.Cells.Columns.Count
    If Cls Then
        WsWb(Ws).Close
    End If
End If
MaxCol = C
End Property

Property Get MaxRow&()
Static R&
If R = 0 Then
    Dim Ws As Worksheet
    Set Ws = ActiveSheet
    Dim Cls As Boolean
    If IsNothing(Ws) Then
        Set Ws = NewWs
        Cls = True
    End If
    R = Ws.Cells.Rows.Count
    If Cls Then
        WsWb(Ws).Close
    End If
End If
MaxRow = R
End Property

