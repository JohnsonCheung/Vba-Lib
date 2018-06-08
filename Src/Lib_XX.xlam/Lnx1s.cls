VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lnx1s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'Private B_Ay() As Lnx1
'Function FmtLy(QQVblStr$) As String()
'Dim O$(), J%
'For J = 0 To U
'    VbAy.Push O, B_Ay(J).Fmt(QQVblStr)
'Next
'FmtLy = O
'End Function
'Property Get Ly() As String()
'Dim O$(), J%
'For J = 0 To U
'    VbAy.Push O, B_Ay(J).Lin
'Next
'Ly = O
'End Property
'Property Get LnxAy() As Lnx1()
'LnxAy = B_Ay
'End Property
'Property Get Dic() As Dictionary
'Set Dic = Vb.Ly(Me.Ly).Dic
'End Property
'Property Get LnxsAy(LinPfx$) As Lnx1s()
'Dim O() As Lnx1s, J&, M As New Lnx1s
'For J = 0 To U
'    Dim Lin$
'    Lin = B_Ay(J).Lin
'    If HasPfx(Lin, "==") Then
'        If M.N > 0 Then
'            PushObj O, M
'            Set M = New Lnx1s
'        End If
'    Else
'        M.Push B_Ay(J)
'    End If
'Next
'LnxsAy = O
'End Property
'Function ValidatePfxAndDupT1(Pfx$) As LnxsRslt
'Stop
'End Function
'Friend Property Get Init(Ay() As Lnx1) As Lnx1s
'B_Ay = Ay
'Set Init = Me
'End Property
'
'Property Get RmvDDRmk() As Lnx1s
'Dim J%, O As New Lnx1s, M As Lnx1
'For J = 0 To U
'    Set M = B_Ay(J).RmvDDRmk
'    If Not M.IsBlankLin Then
'        O.Push M
'    End If
'Next
'Set RmvDDRmk = O
'End Property
'
'Sub Push(A As Lnx1)
'PushObj B_Ay, A
'End Sub
'
'Property Get N&()
'N = Sz(B_Ay)
'End Property
'
'Property Get U&()
'U = N - 1
'End Property
'
'Property Get ToStr$()
'Dim O$(), J%
'For J = 0 To U
'    VbAy.Push O, B_Ay(J).ToStr
'Next
'ToStr = Tag("Lnxs", JnCrLf(O))
'End Property
'Private Sub ZZ_LnxsAy()
'
'End Sub
'
'Sub Tst()
'ZZ_LnxsAy
'End Sub
