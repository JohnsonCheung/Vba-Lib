Attribute VB_Name = "M_Tmp"
Option Explicit
Function TmpFb$(Optional Fdr$, Optional Fnn$)
TmpFb = TmpFfn(".accdb", Fdr, Fnn)
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpFx$(Optional Fdr$, Optional Fnn$)
TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
End Function
Function TmpFxa$(Optional Fdr$, Optional Fnn$)
TmpFxa = TmpFfn(".xlam", Fdr, Fnn)
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function
Private Sub PthEns(P$)
If Not Fso.FolderExists(P) Then MkDir P
End Sub

Function TmpPth$(Optional Fdr$)
Dim X$
   If Fdr <> "" Then
       X = Fdr & "\"
   End If
Dim O$
   O = TmpPthFix & X:   PthEns O
   O = O & TmpNm & "\": PthEns O
   PthEns O
TmpPth = O
End Function

Sub TmpPthBrw()
Shell "Explorer """ & TmpPth & """", vbMaximizedFocus
End Sub

Function TmpPthFix$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthFix = X
End Function
