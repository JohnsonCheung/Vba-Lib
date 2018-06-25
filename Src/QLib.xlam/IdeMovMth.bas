Attribute VB_Name = "IdeMovMth"
Option Explicit

Function CurTarMd() As CodeModule
With CurVbe
   If .CodePanes.Count <> 2 Then Exit Function
   Dim M1 As CodeModule: Set M1 = .CodePanes(1).CodeModule
   Dim M2 As CodeModule: Set M2 = .CodePanes(2).CodeModule
   Dim M As CodeModule: Set M = CurMd
   Dim IsM1Tar As Boolean: IsM1Tar = M1 <> M And M2 = M
   Dim IsM2Tar As Boolean: IsM2Tar = M2 <> M And M1 = M
   If Not (IsM1Tar Xor IsM2Tar) Then Stop
   If IsM1Tar Then Set CurTarMd = M1: Exit Function
   If IsM2Tar Then Set CurTarMd = M2: Exit Function
End With
End Function

Function IsOnlyTwoCdPne() As Boolean
IsOnlyTwoCdPne = CurVbe.CodePanes.Count = 2
End Function

Sub MovAllMth()

End Sub

Sub MovMth()

End Sub

Sub CurTarMd__Tst()
Debug.Print MdNm(CurTarMd)
End Sub
