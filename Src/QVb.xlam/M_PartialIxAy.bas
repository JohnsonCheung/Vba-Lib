Attribute VB_Name = "M_PartialIxAy"
Option Explicit

Function PartialIxAy_CompleteIxAy(PartialIxAy&(), U&) As Long()
'Des:Make a complete-IxAy-of-U by partialIxAy
'Des:A complete-IxAy-Of-U is defined as
'Des:it has (U+1)-elements,
'Des:it does not have dup element
'Des:it has all element of value between 0 and U
Ass IxAy_IsParitial_of_0toU(PartialIxAy, U)
Dim I&(): I = SeqOfLng(0, U)
PartialIxAy_CompleteIxAy = AyAddAp(PartialIxAy, AyMinus(I, PartialIxAy))
End Function
