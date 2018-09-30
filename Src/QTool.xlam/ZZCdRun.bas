Attribute VB_Name = "ZZCdRun"
Sub Mov_Pfx_Dic_To_VbDic()
MthMov DMth("QTool.AX.DicAdd"), Md("VbDic") 'QTool.AX.DicAdd:Fun.
MthMov DMth("QTool.AX.DicClone"), Md("VbDic") 'QTool.AX.DicClone:Fun.
MthMov DMth("QTool.AX.DicCmp"), Md("VbDic") 'QTool.AX.DicCmp:Fun.
MthMov DMth("QTool.AX.DicHasAllKeyIsNm"), Md("VbDic") 'QTool.AX.DicHasAllKeyIsNm:Fun.
MthMov DMth("QTool.AX.DicHasAllValIsStr"), Md("VbDic") 'QTool.AX.DicHasAllValIsStr:Fun.
MthMov DMth("QTool.AX.DicIsEq"), Md("VbDic") 'QTool.AX.DicIsEq:Fun.
MthMov DMth("QTool.AX.DicMinus"), Md("VbDic") 'QTool.AX.DicMinus:Fun.
MthMov DMth("QTool.AX.DicS1S2Itr"), Md("VbDic") 'QTool.AX.DicS1S2Itr:Fun.
MthMov DMth("QTool.AX.DicS1S2Ay"), Md("VbDic") 'QTool.AX.DicS1S2Ay:Fun.
MthMov DMth("QTool.AX.DicSrt"), Md("VbDic") 'QTool.AX.DicSrt:Fun.
MthMov DMth("QTool.AX.DicWb"), Md("VbDic") 'QTool.AX.DicWb:Fun.
MthMov DMth("QTool.AX.DicAy_Mge"), Md("VbDic") 'QTool.AX.DicAy_Mge:Fun.
MthMov DMth("QTool.AX.DicPush"), Md("VbDic") 'QTool.AX.DicPush:Sub.
MthMov DMth("QTool.AX.DicIsEmp"), Md("VbDic") 'QTool.AX.DicIsEmp:Fun.
MthMov DMth("QTool.AX.DicMap"), Md("VbDic") 'QTool.AX.DicMap:Fun.
MthMov DMth("QTool.AX.DicHasStrKy"), Md("VbDic") 'QTool.AX.DicHasStrKy:Fun.
MthMov DMth("QTool.AX.DicHasStrKy1"), Md("VbDic") 'QTool.AX.DicHasStrKy1:Fun.
MthMov DMth("QTool.AX.DicAddKeyPfx"), Md("VbDic") 'QTool.AX.DicAddKeyPfx:Fun.
MthMov DMth("QTool.AX.DicTyBrw"), Md("VbDic") 'QTool.AX.DicTyBrw:Sub.
MthMov DMth("QTool.AX.DicTy"), Md("VbDic") 'QTool.AX.DicTy:Fun.
MthMov DMth("QTool.AX.DicWsBrw"), Md("VbDic") 'QTool.AX.DicWsBrw:Sub.
MthMov DMth("QTool.AX.DicBrw"), Md("VbDic") 'QTool.AX.DicBrw:Sub.
MthMov DMth("QTool.AX.DicLy"), Md("VbDic") 'QTool.AX.DicLy:Fun.
MthMov DMth("QTool.AX.DicWs"), Md("VbDic") 'QTool.AX.DicWs:Fun.
MthMov DMth("QTool.AX.DicStrKy"), Md("VbDic") 'QTool.AX.DicStrKy:Fun.
MthMov DMth("QTool.AX.DicMaxValSz"), Md("VbDic") 'QTool.AX.DicMaxValSz:Fun.
MthMov DMth("QTool.AX.DicAyAdd"), Md("VbDic") 'QTool.AX.DicAyAdd:Fun.
MthMov DMth("QTool.IdePjMge.DicA_RmvMth"), Md("VbDic") 'QTool.IdePjMge.DicA_RmvMth:Fun.
End Sub
Sub Mov_Pj_QTool_Pfx_Has_To_VbStrHas()
MthMov DMth("QTool.AX.HasPfx"), Md("QTool.VbStrHas") 'QTool.AX.HasPfx:Fun.
MthMov DMth("QTool.AX.HasPfxS"), Md("QTool.VbStrHas") 'QTool.AX.HasPfxS:Fun.
MthMov DMth("QTool.AX.HasSubStr"), Md("QTool.VbStrHas") 'QTool.AX.HasSubStr:Fun.
MthMov DMth("QTool.AX.HasPfxAy"), Md("QTool.VbStrHas") 'QTool.AX.HasPfxAy:Fun.
MthMov DMth("QTool.AX.HasSpc"), Md("QTool.VbStrHas") 'QTool.AX.HasSpc:Fun.
MthMov DMth("QTool.IdeBtn.HasBar"), Md("QTool.VbStrHas") 'QTool.IdeBtn.HasBar:Fun.
End Sub
Sub Mov_Pj_QTool_Pfx_Ay_To_VbAy()
MthMov DMth("QTool.AX.ZZ_AyIns"), Md("QTool.VbAy") 'QTool.AX.ZZ_AyIns:Sub.
MthMov DMth("QTool.AX.ZZ_AyabWs"), Md("QTool.VbAy") 'QTool.AX.ZZ_AyabWs:Sub.Prv
MthMov DMth("QTool.AX.ZZ_AyGpCntFmt"), Md("QTool.VbAy") 'QTool.AX.ZZ_AyGpCntFmt:Sub.
MthMov DMth("QTool.AX.AyAddFunCol"), Md("QTool.VbAy") 'QTool.AX.AyAddFunCol:Fun.
MthMov DMth("QTool.IdeMthDr.Z_AyInsAy"), Md("QTool.VbAy") 'QTool.IdeMthDr.Z_AyInsAy:Sub.Prv
MthMov DMth("QTool.IdeMthDr.ZZ_AyReSzAt"), Md("QTool.VbAy") 'QTool.IdeMthDr.ZZ_AyReSzAt:Sub.Prv
MthMov DMth("QTool.VbAyDic.Z_Aydic_to_KeyCntMulItmCol_Dry"), Md("QTool.VbAy") 'QTool.VbAyDic.Z_Aydic_to_KeyCntMulItmCol_Dry:Sub.Prv
MthMov DMth("QTool.VbAyDic.Aydic_to_KeyCntMulItmColDry"), Md("QTool.VbAy") 'QTool.VbAyDic.Aydic_to_KeyCntMulItmColDry:Fun.
End Sub
