Attribute VB_Name = "M_Lnx"
Option Explicit

Property Get LnxRmvDDRmk(A As Lnx) As Lnx
Set LnxRmvDDRmk = Lnx(LinRmvDDRmk(A.Lin), A.Lx)
End Property
