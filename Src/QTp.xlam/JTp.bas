Attribute VB_Name = "JTp"

Property Get ABC(Lin) As ABC
Dim O As New ABC
Set ABC = O.Init(Lin)
End Property

Property Get DDLines(Ly$()) As DDLines
Dim O As New DDLines
Set DDLines = O.Init(Ly)
End Property

Property Get Gp(A() As Lnx) As Gp
Dim O As New Gp
Set Gp = O.Init(A)
End Property

Property Get LABCAyRslt(A() As LABC, ErLy$()) As LABCAyRslt
Dim O As New LABCAyRslt
Set LABCAyRslt = O.Init(A, ErLy)
End Property

Property Get Lnx(Lin, Lx%) As Lnx
Dim O As New Lnx
Set Lnx = O.Init(Lin, Lx)
End Property

Property Get LyRslt(Ly$(), ErLy$()) As LyRslt
Dim O As New LyRslt
Set LyRslt = O.Init(Ly, ErLy)
End Property
