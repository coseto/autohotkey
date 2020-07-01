#NoEnv
StrLeft2Sub(SearchString, Needle )
{
	StringGetPos, varPos, SearchString, %Needle%
	if  errorlevel
	{ 
		return ""
	}
	stringleft varReturn, SearchString, %varPos%
	return %varReturn%
}

StrMid2Sub(varString,subString1,subString2)
{
	StringGetPos, varPos, varString, %subString1%
	if errorlevel
	{
		return ""
	}
	varLen := strlen(subString1)
	varTemp := substr(varString,varPos+varLen+1)
	ifinstring varTemp,%subString2%
	  varTemp := StrLeft2Sub(varTemp,subString2)
	return varTemp
}
StrRight2Sub(varString,subString)
{
	StringGetPos varPos, varString, %subString% , R1
	stringleft varTemp,varString,%varPos%
	varLen := strlen(varTemp)
	varLen := strlen(varString) - varLen - strlen(subString)
	stringright varReturn,varString,%varLen%
	return %varReturn%
}

  sPath := "de_0617_投注计算方式v74.xlsx"
  wPath := 1.抢庄牛牛
  oSheet := ComObjCreate("Excel.Application")
  oSheet.Workbooks.Open(sPath)  ; 開啟已存在的Excel檔案\
  oSheet.Sheets("1.抢庄牛牛").Select
  oSheet.Visible := True


;轉換牌型=>數字
StringReplace, Clipboard, Clipboard,wuniu,1, All
StringReplace, Clipboard, Clipboard,niu1,2, All
StringReplace, Clipboard, Clipboard,niu2,3, All
StringReplace, Clipboard, Clipboard,niu3,4, All
StringReplace, Clipboard, Clipboard,niu4,5, All
StringReplace, Clipboard, Clipboard,niu5,6, All
StringReplace, Clipboard, Clipboard,niu6,7, All
StringReplace, Clipboard, Clipboard,niu7,8, All
StringReplace, Clipboard, Clipboard,niu8,9, All
StringReplace, Clipboard, Clipboard,niu9,10, All
StringReplace, Clipboard, Clipboard,niuniu,11, All
StringReplace, Clipboard, Clipboard,sihuaniu,12, All
StringReplace, Clipboard, Clipboard,shunziniu,13, All
StringReplace, Clipboard, Clipboard,shunjinniu,14, All
StringReplace, Clipboard, Clipboard,wuhuaniu,15, All
StringReplace, Clipboard, Clipboard,tonghuaniu,16, All
StringReplace, Clipboard, Clipboard,huluniu,17, All
StringReplace, Clipboard, Clipboard,wuxiaoniu,18, All
;轉換牌型=>數字




dicedbank1 = ":{"is_banker":true,"dispute_rate":
dicedbank2 = ,"is_done":true
bankrate := StrMid2Sub(Clipboard, dicedbank1, dicedbank2) 	;莊家倍率

diceplayers = {"is_banker":
StrReplace(Clipboard, diceplayers,, checkplayers)			;玩家人數
checkplayerss := checkplayers-1

dicedbasebet1 = {"ante":"
dicedbasebet2 = ","admittance":"
basebet := 0.01*StrMid2Sub(Clipboard, dicedbasebet1, dicedbasebet2)	;底注
StringLeft, OutputVar, String, 4

1pcheckbank = "1":{"is_banker":true
2pcheckbank = "2":{"is_banker":true
3pcheckbank = "3":{"is_banker":true
4pcheckbank = "4":{"is_banker":true

z1pRate = 1":{"raise_rate":
z2pRate = 2":{"raise_rate":
z3pRate = 3":{"raise_rate":
z4pRate = 4":{"raise_rate":
zratedone = ,"is_done":true,

z1p_deckfight = 1":{"cards":(.*),"is_have_bull"(.*),"type":"(.*)","odds":(.*)"2":{"cards ;取牌型字串
z2p_deckfight = 2":{"cards":(.*),"is_have_bull"(.*),"type":"(.*)","odds":(.*)"3":{"cards ;取牌型字串
z3p_deckfight = 3":{"cards":(.*),"is_have_bull"(.*),"type":"(.*)","odds":(.*)"4":{"cards ;取牌型字串
z4p_deckfight = 4":{"cards":(.*),"is_have_bull"(.*),"type":"(.*)","odds":(.*)},"settle	;取牌型字串

z1p_deck = "1(.*)odds":(.*),"is_done":true,"fight_time":(.*),"2":{"cards":
z2p_deck = "2(.*)odds":(.*),"is_done":true,"fight_time":(.*),"3":{"cards":
z3p_deck = "3(.*)odds":(.*),"is_done":true,"fight_time":(.*),"4":{"cards":
z4p_deck = "4(.*)odds":(.*),"is_done":true,"fight_time":(.*)},"settle":

;確定莊家&填寫閒家倍率
IfInString, Clipboard, %1pcheckbank%
{
	Gosub, bank1go
    return
}
else
IfInString, Clipboard, %2pcheckbank%
{
	Gosub, bank2go
    return
}
IfInString, Clipboard, %3pcheckbank%
{
	Gosub, bank3go
    return
}
IfInString, Clipboard, %4pcheckbank%
{
	Gosub, bank4go
    return
}
;四家牌型比大小



	bank1go:
		zArate := StrMid2Sub(Clipboard, z2pRate, zratedone)
		zBrate := StrMid2Sub(Clipboard, z3pRate, zratedone)
		zCrate := StrMid2Sub(Clipboard, z4pRate, zratedone)
		RegExMatch(Clipboard, z1p_deck, z1p_drate)  ; 莊家牌型倍數z1p_drate2.
		RegExMatch(Clipboard, z2p_deck, z2p_drate)  ; 閒家A牌型倍數z2p_drate2.
		RegExMatch(Clipboard, z3p_deck, z3p_drate)  ; 閒家B牌型倍數z3p_drate2.
		RegExMatch(Clipboard, z4p_deck, z4p_drate)  ; 閒家C牌型倍數z4p_drate2.
		RegExMatch(Clipboard, z1p_deckfight, z1p_pk)  ; 莊家牌型
		RegExMatch(Clipboard, z2p_deckfight, z2p_pk)  ; 閒家A牌型
		RegExMatch(Clipboard, z3p_deckfight, z3p_pk)  ; 閒家B牌型
		RegExMatch(Clipboard, z4p_deckfight, z4p_pk)  ; 閒家C牌型
		msgbox go1
		;msgbox % z2p_pk1
		;msgbox % z2p_pk2
		;msgbox % z2p_pk3
		;msgbox % z2p_pk4
		gosub next
	Return

	bank2go:
		zArate := StrMid2Sub(Clipboard, z1pRate, zratedone)
		zBrate := StrMid2Sub(Clipboard, z3pRate, zratedone)
		zCrate := StrMid2Sub(Clipboard, z4pRate, zratedone)
		RegExMatch(Clipboard, z1p_deck, z2p_drate)  ; 閒家A牌型倍數z2p_drate2.
		RegExMatch(Clipboard, z2p_deck, z1p_drate)  ; 莊家牌型倍數z1p_drate2.
		RegExMatch(Clipboard, z3p_deck, z3p_drate)  ; 閒家B牌型倍數z3p_drate2.
		RegExMatch(Clipboard, z4p_deck, z4p_drate)  ; 閒家C牌型倍數z4p_drate2.
		RegExMatch(Clipboard, z1p_deckfight, z2p_pk)  ; 閒家A牌型
		RegExMatch(Clipboard, z2p_deckfight, z1p_pk)  ; 莊家牌型
		RegExMatch(Clipboard, z3p_deckfight, z3p_pk)  ; 閒家B牌型
		RegExMatch(Clipboard, z4p_deckfight, z4p_pk)  ; 閒家C牌型
		msgbox go2
		;msgbox % z1p_pk3
		;msgbox % z2p_pk3
		;msgbox % z3p_pk3
		;msgbox % z4p_pk3
		gosub next
	Return
	bank3go:
		zArate := StrMid2Sub(Clipboard, z1pRate, zratedone)
		zBrate := StrMid2Sub(Clipboard, z2pRate, zratedone)
		zCrate := StrMid2Sub(Clipboard, z4pRate, zratedone)
		RegExMatch(Clipboard, z1p_deck, z2p_drate)  ; 閒家A牌型倍數z2p_drate2.
		RegExMatch(Clipboard, z2p_deck, z3p_drate)  ; 閒家B牌型倍數z3p_drate2.
		RegExMatch(Clipboard, z3p_deck, z1p_drate)  ; 莊家牌型倍數z1p_drate2.
		RegExMatch(Clipboard, z4p_deck, z4p_drate)  ; 閒家C牌型倍數z4p_drate2.
		RegExMatch(Clipboard, z1p_deckfight, z2p_pk)  ; 閒家A牌型
		RegExMatch(Clipboard, z2p_deckfight, z3p_pk)  ; 閒家B牌型
		RegExMatch(Clipboard, z3p_deckfight, z1p_pk)  ; 莊家牌型
		RegExMatch(Clipboard, z4p_deckfight, z4p_pk)  ; 閒家C牌型
		msgbox go3
		;msgbox % z2p_pk1
		;msgbox % z2p_pk2
		;msgbox % z2p_pk3
		;msgbox % z2p_pk4
		gosub next
	Return
	bank4go:
		zArate := StrMid2Sub(Clipboard, z1pRate, zratedone)
		zBrate := StrMid2Sub(Clipboard, z2pRate, zratedone)
		zCrate := StrMid2Sub(Clipboard, z3pRate, zratedone)
		RegExMatch(Clipboard, z1p_deck, z2p_drate)  ; 閒家A牌型倍數z2p_drate2.
		RegExMatch(Clipboard, z2p_deck, z3p_drate)  ; 閒家B牌型倍數z3p_drate2.
		RegExMatch(Clipboard, z3p_deck, z4p_drate)  ; 閒家C牌型倍數z4p_drate2.
		RegExMatch(Clipboard, z4p_deck, z1p_drate)  ; 莊家牌型倍數z1p_drate2.
		RegExMatch(Clipboard, z1p_deckfight, z2p_pk)  ; 閒家A牌型
		RegExMatch(Clipboard, z2p_deckfight, z3p_pk)  ; 閒家B牌型
		RegExMatch(Clipboard, z3p_deckfight, z4p_pk)  ; 閒家C牌型
		RegExMatch(Clipboard, z4p_deckfight, z1p_pk)  ; 莊家牌型
		msgbox go4
		;msgbox, %Clipboard%
		;msgbox, %z1p_pk2%
		;msgbox, %z2p_pk2%
		;msgbox, %z3p_pk2%
		;msgbox, %z4p_pk2%
		gosub next
	Return
next:
ztax1 = "tax":
ztax2 = },"bet_limit
tax := StrMid2Sub(Clipboard, ztax1, ztax2)



oSheet.Range("C3").Value := bankrate
oSheet.Range("C4").Value := checkplayerss
oSheet.Range("C5").Value := basebet
oSheet.Range("C6").Value := zArate
oSheet.Range("C7").Value := zBrate
oSheet.Range("C8").Value := zCrate
oSheet.Range("C10").Value := z1p_drate2
oSheet.Range("C11").Value := z2p_drate2
oSheet.Range("C12").Value := z3p_drate2
oSheet.Range("C13").Value := z4p_drate2
oSheet.Range("C14").Value := tax
oSheet.Range("F7").Value := z1p_pk3
oSheet.Range("I7").Value := z2p_pk3
oSheet.Range("L7").Value := z3p_pk3
oSheet.Range("O7").Value := z4p_pk3

/*

;牌型ok
;to do判斷大小邏輯

LinesFromExcel := Clipboard

Excelceny := []
For Each, Line In StrSplit(LinesFromExcel, ",", ",")
   Excelceny[A_Index] := StrReplace(RegExReplace(Line, "(.*?\t){7}(.+?)\t.*","$2"), ",", ".")

AP_Min := StrReplace(Format("{:.0f}", Min(Excelceny*)), ".", ",") ; Variadic Function Call
AP_Max := StrReplace(Format("{:.0f}", Max(Excelceny*)), ".", ",")
;MsgBox, Min: %AP_Min% - Max: %AP_Max%
*/