аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџTotalWD
		.Font.Bold = True
	End With
	With	sh.Cell(sr, 9)
		.Value = ri.TotalIN - ri.TotalWD
		.Font.Bold = True
	End With
	Call DrawIN(sh, ri, sr + 1)
	Call DrawWD(sh, ri, sr + 1)
	DrawItem = sz
End Function 

Function DrawTotal(sr)
	Dim sh
	Dim sz
	Set sh = ActiveSheet
	sz = GetMax(1, rp.CountIN + 1, rp.CountWD + 1)
	With sh.Cell(sr, 2)
		.Value = "Чрурыюь"
		.Font.Bold = True
	End With
	sh.Cell(sr, 3).Value = "Тёќюую"
	With sh.Range(3, sr, 4, sr)
		.MergeCells = True
		.SetBorder acBrdOutline
	End With
	sh.Cell(sr, 6).Value = "Тёќюую"
	With sh.Range(6, sr, 7, sr)
		.MergeCells = True
		.SetBorder acBrdOutline
	End With
	With sh.Cell(sr, 5)
		.Value = rp.TotalIN
		.Font.Bold = True
	End With
	With sh.Cell(sr, 8)
		.Value = rp.TotalWD
		.Font.Bold = True
	End With
	With sh.Cell(sr, 9)
		.Value = rp.TotalIN - rp.TotalWD
		.Font.Bold = True
	End With
	sh.Range(1, sr, COLS, sr + sz - 1).SetBorder acBrdOutline	+ acBrdInsideVert
	Call DrawTotalIN(sh, sr + 1)
	Call DrawTotalWD(sh, sr + 1)
	DrawTotal = sz
End Function 

Function GetSiteInfo()
	GetSiteInfo = 0
	Dim st
	Set st = Workarea.Site
	If st.Kind <> acAgent Then Exit Function
	If st.Type <> 1 And st.Type <> 2 Then Exit Function
	GetSiteInfo = st.ID
End Function 

Sub DrawSiteError
	Dim sh
	Set sh = ActiveSheet
	sh.Rows = 1
	sh.Cell(1,1).Value = "арёїхђэћх ышёђъш ьюцэю ёђ№юшђќ ђюыќъю"
	sh.Cell(2,1).Value = "яю яюф№рчфхыхэшў шыш я№хфя№шџђшў"
	sh.Range(1, 1, 1, 2).ForeColor = vbRed
	sh.SetRealSize
	sh.AutoFit acColumns
End Sub

Sub ShtBook_OnLoad
	Dim row
	Dim i
	Dim sh
	Dim br
	Dim wc
	Dim tt
	Set wc = WaitCursor
	SetDefaultFont "Verdana", 11
	Set sh = ActiveSheet
	sh.Rows = 1
	sh.DisplayHeadings = False
	sh.DisplayGrid = False
	Set rp = Workarea.CreateReport("PRepAgentBill")
	rp.AgID = GetSiteInfo()
	rp.SortMode = 1
	If rp.AgID = 0 Then 
		DrawSiteError
		Exit Sub ' тёх ѓцх ѓёђрэютыхэю
	End If
	rp.SetOrgMode ' №хцшь №рсюђћ "Фыџ тёхѕ"
	rp.Build
	row = DrawHeader()
	br = row
	For i=1 To rp.Count
		row = row + DrawItem(row, i)
	Next
	sh.Range(1, br, 1, row-1).Alignment = acCenter
	row = row + DrawTotal(row)
	' тћ№ртэштрэшх ш єю№ьрђћ фыџ ёѓьь
	sh.Range(5, br, 5, row-1).Alignment = acRight
	sh.Range(8, br, 8, row-1).Alignment = acRight
	sh.Range(9, br, 9, row-1).Alignment = acRight
	' тћ№ртэштрэшх фыџ ьхёџіхт
	sh.Range(3, br, 3, row-1).Alignment = acCenter
	sh.Range(6, br, 6, row-1).Alignment = acCenter

	tt = rp.TotalIN - rp.TotalWD
	sh.Cell(row, 1).Value = "Фю ёяырђш: " & SpellMoney2(tt, SPELL_FMT, 1, 0)
	sh.Range(1, row, COLS, row).MergeCells = True
	
	row = row + 2
	sh.Cell(row, 1).Value = """_____"" ____________ 20___ №."
	sh.Range(1, row, 5, row).MergeCells = True

	sh.Cell(row, 6).Value = "Сѓѕурыђх№ ________________ "
	sh.Cell(row, 6).Alignment = acRight
	sh.Range(6, row, COLS, row).MergeCells = True


	sh.SetRealSize
	sh.AutoFit acColumns + acRows
End Sub
<     ш    м           Ь Verdana                         ЭЭ      џџџџ                 џџ 
 CSheetPageSheet1Ышёђ 1         61    7                                     W      џџ  CRow џџ  CCell6    џџџџ                      d                <     ш                Ь Arial ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ <     ш                Ь Arial ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ <     ш                Ь Arial ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ <     ш                Ь Arial ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ <     ш                Ь Arial ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ <     ш                Ь Arial ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ    2   2       <     ш    м       џ   Ь Verdana ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ <     ш   м           Ь Verdana ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ <     ш   №           Ь Verdana ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ  <     ш    №           Ь Verdana ЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭЭ   
 џџџџ                             	 џџ                                   џџ                                џџ                                 џџ                                џџџџ                              џџџџ                               џџџџ                              џџџџ                              ррр                                  ррр                               _ _ _ _ _ _ _                       џџџџ                         џџџџ                         џџџџ                         џџџџ                       	      џџ	 x    >7@0EC=:>20- ?;0BV6=0  2V4><VABL  !  _ _ _ _ _ _ _     70  8?5=L  2 0 0 3   R o o t   E n t r y                                               џџџџџџџџ                               А5яШФ         C o n t e n t s                                                  џџџџ   џџџџ                                       Ъ        S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        e                                                                          џџџџџџџџџџџџ                                                                        ўџџџўџџџ§џџџўџџџўџџџ                         џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ    рђљOhЋ +'Гй   рђљOhЋ +'Гй0   5        x                   Ќ      И      Ф      а   	   м   
   ш      є                       (     у        №рёїхђэћх_ышёђъш.ash                                                            173 @   0Р­   @   а0шШФ@   p sЙУ      Ръіхэђ 7.0                                               џњOption Explicit

Const COLS = 8
Dim rp

Function DrawHeader
	Dim HEADER 
   HEADER = _
		"1:ЯВС, яюёрфр"	& vbLf & _ џЃOption Explicit

Const COLS = 9
Const SPELL_FMT = "0|у№штхэќ|у№штэџ|у№штэГ|0|ъюяГщюъ|ъюяГщър|ъюяГщъш"
Dim rp

Function GetAgPlace(ID)
	GetAgPlace = ""
	If ID = 0 Then Exit Function
	Dim msc
	Set msc = Workarea.Misc(ID)
	If msc Is Nothing Then Exit Function
	GetAgPlace = msc.Name
End Function


Function GetAgIXTag(ID)
	Dim s
	Dim e
	s = "Ючэрър фюѕюфѓ: 01"
	If ID <> 0 Then
     e = Workarea.Enum(2).ItemNameID(ID)
		s = "Ючэрър фюѕюфѓ: " & e
	End If
	GetAgIXTag = s
End Function

Function GetCodeText(cd)
  	Dim s
	Dim id
	s = cd & " ("
	id = Workarea.GetCodeID(cd)
	If id <> 0 Then
		s = s + Workarea.Code(id).Tag
	End If
	s = s + ")"
	GetCodeText = s
End Function 

Function DrawHeader
	Dim HEADER 
   HEADER = _
		"1:Й"	& vbLf & _ 
		"1:ЯВС, яюёрфр"	& vbLf & _ 
     "1:Эр№рѕютрэю"	& vbLf & _ 
	     "2:ЬГё."	& vbLf & _ 
	     "2:Ъюф"	& vbLf & _ 
	     "2:бѓьр"	& vbLf & _ 
   "1:гђ№шьрэю"   	 	& vbLf & _ 
	     "2:ЬГё."	& vbLf & _ 
	     "2:Ъюф"	& vbLf & _ 
	     "2:бѓьр"	& vbLf & _ 
	 "1:Фю ёяырђш"  	& vbLf & _ 
		"2:ЯГфяшё" 
	Dim sh
	Set sh = ActiveSheet
	With sh.Cell(1, 1)
		.Value = "Т ърёѓ фю ёяырђш ч ___________ яю  ___________"
		.Alignment = acRight
	End With
	With sh.Cell(2, 1)
		.Value = "Тшфрђъютшщ ърёютшщ ю№фх№ Й _____ тГф ""___""________20__№."
		.Alignment = acRight
	End With
	sh.Cell(4, 1).Value = "ахъђю№ _______________"
	sh.Range(1, 4, 4, 4).MergeCells = True
	With sh.Cell(4, 5)
		.Value = "Уюыютэшщ сѓѕурыђх№ _______________"
		.Alignment = acRight
	End With
	sh.Range(5, 4, COLS, 4).MergeCells = True
	With sh.Cell(6, 1)
		.Value = "аюч№рѕѓэъютр-яырђГцэр тГфюьГёђќ Й _______  чр "  & _
            FormatDate2(DateSerial(rp.Period.Year, rp.Period.Month, 1), "(ua)mmmm yyyy №.")
		.Font.Bold = True
	End With
	sh.Range(1, 1, COLS, 1).MergeCells = True
	sh.Range(1, 2, COLS, 2).MergeCells = True
	sh.Range(1, 6, COLS, 6).MergeCells = True
	Call sh.MakeHeader(8, 1, HEADER)
	DrawHeader = 10
End Function

Function GetMax(v1, v2, v3)
	Dim v
	v = v1
	If v2 > v Then v = v2
	If v3 > v Then v = v3
	GetMax = v
End Function

Sub DrawIN(sh, ri, sr)
	Dim ra
	Dim i
	For i=1 To ri.CountIN
		Set ra = ri.ItemIN(i)
		With ra
			sh.Cell(sr + i - 1, 3).Value = .Month
			sh.Cell(sr + i - 1, 4).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 5).Value = .Sum
		End With
	Next
End Sub

Sub DrawTotalIN(sh, sr)
	Dim ra
	Dim i
	For i=1 To rp.CountIN
		Set ra = rp.ItemIN(i)
		With ra
			sh.Cell(sr + i - 1, 4).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 5).Value = .Sum
		End With
	Next
End Sub

Sub DrawTotalWD(sh, sr)
	Dim ra
	Dim i
	For i=1 To rp.CountWD
		Set ra = rp.ItemWD(i)
		With ra
			sh.Cell(sr + i - 1, 7).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 8).Value = .Sum
		End With
	Next
End Sub


Sub DrawWD(sh, ri, sr)
	Dim ra
	Dim i
	For i=1 To ri.CountWD
		Set ra = ri.ItemWD(i)
		With ra
			sh.Cell(sr + i - 1, 6).Value = .Month
			sh.Cell(sr + i - 1, 7).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 8).Value = .Sum
		End With
	Next
End Sub


Function DrawItem(sr, i)
	Dim sh
	Dim ri
	Dim ag
	Dim sz
	Set sh = ActiveSheet
	Set ri = rp.Item(i)
	Set ag = Workarea.Agent(ri.AgID)
	sz = GetMax(5, ri.CountIN + 1, ri.CountWD + 1)
	sh.Cell(sr, 1).Value = CStr(i)
	sh.Cell(sr, 2).Value = ag.Name
	sh.Cell(sr + 1, 2).Value = "врс. Й " & ag.TabNo
	sh.Cell(sr + 2, 2).Value = "Iэф. Й " & ag.VatNo
	sh.Cell(sr + 3, 2).Value = GetAgIXTag(ag.Prop("Ючэрър фюѕюфѓ"))
	With sh.Cell(sr + 4, 2)
		.Value = GetAgPlace(ag.Prop("Яюёрфр"))
		.Multiline = True
	End With
	sh.Range(1, sr, COLS, sr + sz-1).SetBorder acBrdOutline	+ acBrdInsideVert
	sh.Cell(sr, 3).Value = "Тёќюую"
	With sh.Range(3, sr, 4, sr)
		.MergeCells = True
		.SetBorder acBrdOutline
	End With
	sh.Cell(sr, 6).Value = "Тёќюую"
	With sh.Range(6, sr, 7, sr)
		.MergeCells = True
		.SetBorder acBrdOutline
	End With
	With sh.Cell(sr, 5)
		.Value = ri.TotalIN
		.Font.Bold = True
	End With
	With sh.Cell(sr, 8)
		.Value = ri.R o o t   E n t r y                                               џџџџџџџџ                               pф^MІФ         C o n t e n t s                                                  џџџџ   џџџџ                                       Ъ        S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        e                                                                          џџџџџџџџџџџџ                                                                        ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ                          џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ    рђљOhЋ +'Гй   рђљOhЋ +'Гй0   5        x                   Ќ      И      Ф      а   	   м   
   ш      є                       (     у        №рёїхђэћх_ышёђъш.ash                                                            174 @   РЩ   @    BNMІФ@   p sЙУ      Ръіхэђ 7.0                                               џњOption Explicit

Const COLS = 8
Dim rp

Function DrawHeader
	Dim HEADER 
   HEADER = _
		"1:ЯВС, яюёрфр"	& vbLf & _ џЃOption Explicit

Const COLS = 9
Const SPELL_FMT = "0|у№штхэќ|у№штэџ|у№штэГ|0|ъюяГщюъ|ъюяГщър|ъюяГщъш"
Dim rp

Function GetAgPlace(ID)
	GetAgPlace = ""
	If ID = 0 Then Exit Function
	Dim msc
	Set msc = Workarea.Misc(ID)
	If msc Is Nothing Then Exit Function
	GetAgPlace = msc.Name
End Function


Function GetAgIXTag(ID)
	Dim s
	Dim e
	s = "Ючэрър фюѕюфѓ: 01"
	If ID <> 0 Then
     e = Workarea.Enum(2).ItemNameID(ID)
		s = "Ючэрър фюѕюфѓ: " & e
	End If
	GetAgIXTag = s
End Function

Function GetCodeText(cd)
  	Dim s
	Dim id
	s = cd & " ("
	id = Workarea.GetCodeID(cd)
	If id <> 0 Then
		s = s + Workarea.Code(id).Tag
	End If
	s = s + ")"
	GetCodeText = s
End Function 

Function DrawHeader
	Dim HEADER 
   HEADER = _
		"1:Й"	& vbLf & _ 
		"1:ЯВС, яюёрфр"	& vbLf & _ 
     "1:Эр№рѕютрэю"	& vbLf & _ 
	     "2:ЬГё."	& vbLf & _ 
	     "2:Ъюф"	& vbLf & _ 
	     "2:бѓьр"	& vbLf & _ 
   "1:гђ№шьрэю"   	 	& vbLf & _ 
	     "2:ЬГё."	& vbLf & _ 
	     "2:Ъюф"	& vbLf & _ 
	     "2:бѓьр"	& vbLf & _ 
	 "1:Фю ёяырђш"  	& vbLf & _ 
		"2:ЯГфяшё" 
	Dim sh
	Set sh = ActiveSheet
	With sh.Cell(1, 1)
		.Value = "Т ърёѓ фю ёяырђш ч ___________ яю  ___________"
		.Alignment = acRight
	End With
	With sh.Cell(2, 1)
		.Value = "Тшфрђъютшщ ърёютшщ ю№фх№ Й _____ тГф ""___""________20__№."
		.Alignment = acRight
	End With
	sh.Cell(4, 1).Value = "ахъђю№ _______________"
	sh.Range(1, 4, 4, 4).MergeCells = True
	With sh.Cell(4, 5)
		.Value = "Уюыютэшщ сѓѕурыђх№ _______________"
		.Alignment = acRight
	End With
	sh.Range(5, 4, COLS, 4).MergeCells = True
	With sh.Cell(6, 1)
		.Value = "аюч№рѕѓэъютр-яырђГцэр тГфюьГёђќ Й _______  чр "  & _
            FormatDate2(DateSerial(rp.Period.Year, rp.Period.Month, 1), "(ua)mmmm yyyy №.")
		.Font.Bold = True
	End With
	sh.Range(1, 1, COLS, 1).MergeCells = True
	sh.Range(1, 2, COLS, 2).MergeCells = True
	sh.Range(1, 6, COLS, 6).MergeCells = True
	Call sh.MakeHeader(8, 1, HEADER)
	DrawHeader = 10
End Function

Function GetMax(v1, v2, v3)
	Dim v
	v = v1
	If v2 > v Then v = v2
	If v3 > v Then v = v3
	GetMax = v
End Function

Sub DrawIN(sh, ri, sr)
	Dim ra
	Dim i
	For i=1 To ri.CountIN
		Set ra = ri.ItemIN(i)
		With ra
			sh.Cell(sr + i - 1, 3).Value = .Month
			sh.Cell(sr + i - 1, 4).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 5).Value = .Sum
		End With
	Next
End Sub

Sub DrawTotalIN(sh, sr)
	Dim ra
	Dim i
	For i=1 To rp.CountIN
		Set ra = rp.ItemIN(i)
		With ra
			sh.Cell(sr + i - 1, 4).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 5).Value = .Sum
		End With
	Next
End Sub

Sub DrawTotalWD(sh, sr)
	Dim ra
	Dim i
	For i=1 To rp.CountWD
		Set ra = rp.ItemWD(i)
		With ra
			sh.Cell(sr + i - 1, 7).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 8).Value = .Sum
		End With
	Next
End Sub


Sub DrawWD(sh, ri, sr)
	Dim ra
	Dim i
	For i=1 To ri.CountWD
		Set ra = ri.ItemWD(i)
		With ra
			sh.Cell(sr + i - 1, 6).Value = .Month
			sh.Cell(sr + i - 1, 7).Value = GetCodeText(.Code)
			sh.Cell(sr + i - 1, 8).Value = .Sum
		End With
	Next
End Sub


Function DrawItem(sr, i)
	Dim sh
	Dim ri
	Dim ag
	Dim sz
	Set sh = ActiveSheet
	Set ri = rp.Item(i)
	Set ag = Workarea.Agent(ri.AgID)
	sz = GetMax(5, ri.CountIN + 1, ri.CountWD + 1)
	sh.Cell(sr, 1).Value = CStr(i)
	sh.Cell(sr, 2).Value = ag.Name
	sh.Cell(sr + 1, 2).Value = "врс. Й " & ag.TabNo
	sh.Cell(sr + 2, 2).Value = "Iэф. Й " & ag.VatNo
	sh.Cell(sr + 3, 2).Value = GetAgIXTag(ag.Prop("Ючэрър фюѕюфѓ"))
	With sh.Cell(sr + 4, 2)
		.Value = GetAgPlace(ag.Prop("Яюёрфр"))
		.Multiline = True
	End With
	sh.Range(1, sr, COLS, sr + sz-1).SetBorder acBrdOutline	+ acBrdInsideVert
	sh.Cell(sr, 3).Value = "Тёќюую"
	With sh.Range(3, sr, 4, sr)
		.MergeCells = True
		.SetBorder acBrdOutline
	End With
	sh.Cell(sr, 6).Value = "Тёќюую"
	With sh.Range(6, sr, 7, sr)
		.MergeCells = True
		.SetBorder acBrdOutline
	End With
	With sh.Cell(sr, 5)
		.Value = ri.TotalIN
		.Font.Bold = True
	End With
	With sh.Cell(sr, 8)
		.Value = ri.