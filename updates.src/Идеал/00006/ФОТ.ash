аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ i = LeftPos To Sh.Columns
		rez = 0
		For j = pos To Sh.Rows - 1
			If IsNumeric(Sh.Cell(j,i).Value) Then _
				rez = rez + CDbl(Sh.Cell(j,i).Value)
		Next
		Sh.Cell(Sh.Rows,i).Value = rez
	Next

	sh.Range(1,pos,sh.columns-1,pos+agCount).SetBorder acBrdGrid	

	Sh.Cell(Sh.Rows,LeftPos-1).Value = "Шђюую: "
	sh.Range(LeftPos,pos,sh.columns-1,pos+agCount+1).SetBorder acBrdGrid	
	sh.Range(LeftPos,pos+1,sh.columns-1,pos+agCount+1).Alignment acRight	

	sh.Range(3,pos+1,3,pos+agCount+1).AutoFit acColumns ' Agent ID

	recalc
End Sub

'=================
Function GetSiteInfo()
	Dim st
	Dim rez	
	rez = 0
	Set st = WorkArea.Site
	If st.Kind <> acAgent Then Exit Function
	If st.Type <> 1 And st.type <> 2 Then Exit Function
	rez = st.ID
	GetSiteInfo	= rez
End Function<     ш                Ь Arial D:\STDs.!!!\Zap\Шфхры\Шфхры       џџџџ                 џџ 
 CSheetPageSheet1Ышёђ 1	   	      61    7                                    $ *   D   F                     -      -      -      -   4      %   #      #   #                   #   =   6   .   +   G   F   +          <            	   џџ  CRow џџ  CCell6     џџ   >   $"  ?>  ?>4@0745;5=8N:   @C??0  !1                      5 џџ   
   !  ?/ ?                    5 џџ      $0<8;8O  . .                     5 џџ   6   45=B8-   D8:0F8>=-   =K9  =><5@                      џџ                           џџ                           џџ                           џџ                           џџ                          5 џџ                           5 џџ                             џџџџ                      ! џџ    2   5AG0AB=K9  A;CG09  ( 1 . 0 2 % )                     ! џџ    2   5B@C4>A?>A>1=>ABL  ( 2 . 9 % )                       џџ    *   5=A8>==K9  D>=4  ( 3 2 % )                       џџ    (   5=A8>==K9  D>=4  ( 4 % )                       џџ    .   $>=4  157@01>B8FK  ( 1 . 9 % )                   	   5 џџ                           5 џџ                             џџџџ                      ! џџ       2 6 6 0   3@=                    ! џџ       2 6 6 0   3@=                    ! џџ       2 6 6 0   3@=                    ! џџ       2 6 6 0   3@=                    ! џџ       2 6 6 0   3@=                    ! џџџџ                      	   5 џџ                           5 џџ                             џџ                        ! џџ     
   1 , 0 2 %                     ! џџ        2 , 9 %                     ! џџ        3 2 %                     ! џџ        4 %                     ! џџ        1 , 9 %                     ! џџџџ                           џџ                            џџ     6   >;5A=8G5=:>  8E08;  .@L528G                      џџ        2 8 9 8 3 0 5 3 9 5                     " џџ                           " џџ                           " џџ                           " џџ                           " џџ                              џџ                            џџ     8   $54G5=:>  =4@59  0;5=B8=>28G                      џџ        2 7 3 7 4 1 3 7 9 3                     " џџ     +йЮ@                    " џџ     ЎGсz-@                    " џџ     ffffffd@                    " џџ                           " џџ     сzЎGa#@                  	         џџџџ      B>3>:                       " џџ     +йЮ@                    " џџ     ЎGсz-@                    " џџ     ffffffd@                    " џџ                                 " џџ     сzЎGa#@                      џџџџ                                 d       d    <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                2   2       <     ш   №           Ь Arial рџђрђќ ярэхыќ шэёђ№ѓьхэђют  <     ш            РРР Ь Arial в ЌВI    b!  ГIЭЭЭЭЭЭЭЭ <     ш            џџџ Ь Arial ш \БI    К$ АБI     <     ш   Ш           Ь Arial                               <     ш               Ь Arial џџџџ                       <     ш              Ь Arial рџђрђќ ярэхыќ шэёђ№ѓьхэђют  <     ш             Ь Arial ЈЕA        ,Ы        џџџџ <     ш    №           Ь Arial ЈЕA        ЄО        џџџџ    џџџџ                               џџџџ                               џџџџ                             џџџџ                                                                                                      Ш                                    Ш                                 t e n t s                                                  џџџџ   џџџџ                                       $       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        U                                                                          џџџџџџџџџџџџ                                                                        	      ўџџџ§џџџўџџџўџџџўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ                          џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   %        x                        Ј      Д      Р   	   Ь   
   и      ф      №      ќ                у        дЮв.ash                                                         525 @   PF/Џ\  @   @у`1њФ@   `јЛ
ІФ      Ръіхэђ 7.0                                                                                                                                                                                                                                                 Ш                                    Ш                                 o n t e n t s                                                  џџџџ   џџџџ                                       р"       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        U                                                                          џџџџџџџџџџџџ                                                R o o t   E n t r y                                               џџџџџџџџ                               аеMo1њФ         C o n t e n t s                                                  џџџџ   џџџџ                                       $       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        U                                                                          џџџџџџџџџџџџ                                                                        	   
   ўџџџџџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ               !   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ"       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   %        x                        Ј      Д      Р   	   Ь   
   и      ф      №      ќ                у        дЮв.ash                                                         526 @   №}uО\  @   рIo1њФ@   `јЛ
ІФ      Ръіхэђ 7.0                                                                                                                                                                                              џOption Explicit

Dim Rp, Sh
Dim IDgroup
Dim arrFot()
Dim pp,mm,yy

Const LeftPos = 4

'=================
Sub ShtBook_OnLoad
	Dim i
	Dim pos
	Dim rez
	Dim ri
	Dim ag
	Dim agCount
	Dim mtdCount
	Dim j
	Dim rs,rs2 ' фыџ чря№юёр
	Dim strSql,db,qd
	Dim sqlAdd1, sqlAdd2, sqlAdd8

	Set Sh = ActiveSheet
	pos = 6 ' Ъюышїхёђтю ёђ№юъ фыџ јряъш

	Sh.Rows = pos

	IDgroup = GetSiteInfo()
	If IDgroup = 0 Then
		MsgBox "Юјшсър. Эхђ яюф№рчфхыхэшџ."
		Exit Sub
	End If

	Sh.Cell(2,1).Value = "дЮв яю яюф№рчфхыхэшў: "&WorkArea.Agent(IDgroup).Name ' Эрчтрэшх яюф№рчфхыхэшџ

'	Header
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT MTD_ID, MTD_NAME, MTD_PERCENT, MTD_MODE, MTD_THRESHOLD FROM MTD_CODES WHERE MTD_TYPE=103 ORDER BY MTD_NAME")
	i = 0
	Sh.Columns = LeftPos + rs.recordcount
	ReDim arrFot(rs.recordcount,4)
	While Not rs.eof	
		arrFot(i,1) = rs.fields("MTD_ID")
		arrFot(i,2) = IIF(Not IsNull(rs.fields("MTD_MODE")),rs.fields("MTD_MODE"),0)
		arrFot(i,3) = IIF(Not IsNull(rs.fields("MTD_PERCENT")),rs.fields("MTD_PERCENT"),0)
		arrFot(i,3) = IIF(Not IsNull(rs.fields("MTD_PERCENT")),rs.fields("MTD_PERCENT"),0)
		arrFot(i,4) = IIF(Not IsNull(rs.fields("MTD_THRESHOLD")),rs.fields("MTD_THRESHOLD"),0)

		Sh.Cell(4,LeftPos+i).Value = rs.fields("MTD_NAME")	
		Sh.Cell(6,LeftPos+i).Value = arrFot(i,3)&"%"
		Sh.Cell(5,LeftPos+i).Value = arrFot(i,4)&" у№э"
		sh.Range(LeftPos+i,3,LeftPos+i,6).SetBorder acBrdOutline	
		i = i + 1
		rs.moveNext
	Wend
	mtdCount = i-1
	sh.Range(LeftPos,4,LeftPos+i,5).AutoFit ' 4 4 8 5
	sh.Range(LeftPos,5,LeftPos+i,6).Alignment acCenter

	Set pp = WorkArea.Period
	yy = pp.year
	mm = pp.month
'	Agents
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT AGENTS.AG_TYPE, AG_TREE.ID FROM AGENTS INNER Join AG_TREE On AGENTS.AG_ID = AG_TREE.ID WHERE P0="&IDgroup&" AND AG_TYPE=3 ORDER BY AGENTS.AG_NAME;")
	i = 1
	agCount = rs.recordCount
	Sh.Rows = agCount + pos
	While Not rs.eof	
		Set ag = WorkArea.Agent(rs.fields("ID"))
		Sh.Cell(i+pos,1).Value = i					' Яю№џфюъютћщ эюьх№
		Sh.Cell(i+pos,2).Value = ag.Name		' Шьџ ёюђ№ѓфэшър
		Sh.Cell(i+pos,3).Value = ag.VatNo		' Шфхэђшєшърішюээћщ ъюф

		' Fill FOT for each AG
		For j = 0 To mtdCount

			sqlAdd1 = ""
			sqlAdd8 = " AND AG_AGRMNT IS NULL "

			If arrFot(j,2) > 0 Then
					If arrFot(j,2) And 8 Then _
						sqlAdd8 = ""
					If arrFot(j,2) And 1 Then _
						sqlAdd1 = " AND (Not AG_DISAB Or AG_DISAB Is Null) "
					If arrFot(j,2) And 2 Then _
						sqlAdd1 = " AND AG_DISAB "
			Else
				' Фыџ тёхѕ
			End If

			strSql = ""+ _
				"Select SUM(P_JOURNAL.JP_SUM) as TOTSUM, P_JOURNAL.AG_ID, PP_MONTH, PP_YEAR, MTD_ID_RIGHT, AG_DISAB, AG_AGRMNT" +_
				" FROM (((((MTD_DEPENDS Left Join P_JOURNAL On MTD_DEPENDS.MTD_ID_LEFT = P_JOURNAL.MTC_ID) Left Join MTD_CODES On MTD_CODES.MTD_ID = P_JOURNAL.MTC_ID) Left Join P_PERIOD On P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD) INNER Join GetFotIds On GetFotIds.FOT_ID = MTD_DEPENDS.MTD_ID_RIGHT) Left Join AG_TREE On P_JOURNAL.AG_ID=AG_TRџOption Explicit

Dim Rp, Sh
Dim IDgroup
Dim arrFot()
Dim pp,mm,yy

Const LeftPos = 4

'=================
Sub ShtBook_OnLoad
	Dim i
	Dim pos
	Dim rez
	Dim ri
	Dim ag
	Dim agCount
	Dim mtdCount
	Dim j
	Dim rs,rs2 ' фыџ чря№юёр
	Dim strSql,db,qd
	Dim sqlAdd1, sqlAdd2, sqlAdd8

	Set Sh = ActiveSheet
	pos = 6 ' Ъюышїхёђтю ёђ№юъ фыџ јряъш

	Sh.Rows = pos

	IDgroup = GetSiteInfo()
	If IDgroup = 0 Then
		MsgBox "Юјшсър. Эхђ яюф№рчфхыхэшџ."
		Exit Sub
	End If

	Sh.Cell(2,1).Value = "дЮв яю яюф№рчфхыхэшў: "&WorkArea.Agent(IDgroup).Name ' Эрчтрэшх яюф№рчфхыхэшџ

'	Header
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT MTD_ID, MTD_NAME, MTD_PERCENT, MTD_MODE, MTD_THRESHOLD FROM MTD_CODES WHERE MTD_TYPE=103 ORDER BY MTD_NAME")
	i = 0
	Sh.Columns = LeftPos + rs.recordcount
	ReDim arrFot(rs.recordcount,4)
	While Not rs.eof	
		arrFot(i,1) = rs.fields("MTD_ID")
		arrFot(i,2) = IIF(Not IsNull(rs.fields("MTD_MODE")),rs.fields("MTD_MODE"),0)
		arrFot(i,3) = IIF(Not IsNull(rs.fields("MTD_PERCENT")),rs.fields("MTD_PERCENT"),0)
		arrFot(i,3) = IIF(Not IsNull(rs.fields("MTD_PERCENT")),rs.fields("MTD_PERCENT"),0)
		arrFot(i,4) = IIF(Not IsNull(rs.fields("MTD_THRESHOLD")),rs.fields("MTD_THRESHOLD"),0)

		Sh.Cell(4,LeftPos+i).Value = rs.fields("MTD_NAME")	
		Sh.Cell(6,LeftPos+i).Value = arrFot(i,3)&"%"
		Sh.Cell(5,LeftPos+i).Value = arrFot(i,4)&" у№э"
		sh.Range(LeftPos+i,3,LeftPos+i,6).SetBorder acBrdOutline	
		i = i + 1
		rs.moveNext
	Wend
	mtdCount = i-1
	sh.Range(LeftPos,4,LeftPos+i,5).AutoFit ' 4 4 8 5
	sh.Range(LeftPos,5,LeftPos+i,6).Alignment acCenter

	Set pp = WorkArea.Period
	yy = pp.year
	mm = pp.month
'	Agents
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT AGENTS.AG_TYPE, AG_TREE.ID FROM AGENTS INNER Join AG_TREE On AGENTS.AG_ID = AG_TREE.ID WHERE P0="&IDgroup&" AND AG_TYPE=3 ORDER BY AGENTS.AG_NAME;")
	i = 1
	agCount = rs.recordCount
	Sh.Rows = agCount + pos
	While Not rs.eof	
		Set ag = WorkArea.Agent(rs.fields("ID"))
		Sh.Cell(i+pos,1).Value = i					' Яю№џфюъютћщ эюьх№
		Sh.Cell(i+pos,2).Value = ag.Name		' Шьџ ёюђ№ѓфэшър
		Sh.Cell(i+pos,3).Value = ag.VatNo		' Шфхэђшєшърішюээћщ ъюф

		' Fill FOT for each AG
		For j = 0 To mtdCount

			sqlAdd1 = ""
			sqlAdd8 = " AND AG_AGRMNT IS NULL "

			If arrFot(j,2) > 0 Then
					If arrFot(j,2) And 8 Then _
						sqlAdd8 = ""
					If arrFot(j,2) And 1 Then _
						sqlAdd1 = " AND (Not AG_DISAB Or AG_DISAB Is Null) "
					If arrFot(j,2) And 2 Then _
						sqlAdd1 = " AND AG_DISAB "
			Else
				' Фыџ тёхѕ
			End If

			strSql = ""+ _
				"Select SUM(P_JOURNAL.JP_SUM) as TOTSUM, P_JOURNAL.AG_ID, PP_MONTH, PP_YEAR, MTD_ID_RIGHT, AG_DISAB, AG_AGRMNT" +_
				" FROM (((((MTD_DEPENDS Left Join P_JOURNAL On MTD_DEPENDS.MTD_ID_LEFT = P_JOURNAL.MTC_ID) Left Join MTD_CODES On MTD_CODES.MTD_ID = P_JOURNAL.MTC_ID) Left Join P_PERIOD On P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD) INNER Join GetFotIds On GetFotIds.FOT_ID = MTD_DEPENDS.MTD_ID_RIGHT) Left Join AG_TREE On P_JOURNAL.AG_ID=AG_TREE.ID) Left Join AGENTS On AGENTS.AG_ID = P_JOURNAL.AG_ID" +_
				" WHERE MTD_TYPE = 101 And P0 = [OrgID] And PP_MONTH=[mm] And PP_YEAR=[yy] And P_JOURNAL.AG_ID = [AgID] And MTD_ID_RIGHT=[FotID] " +_
				sqlAdd1 + sqlAdd8 +_
				" GROUP BY P_JOURNAL.AG_ID, MTD_ID_RIGHT, PP_MONTH, PP_YEAR, AG_DISAB, AG_AGRMNT ;" 

			Set db = workArea.DAODataBase
			Set qd = db.CreateQueryDef("")
			qd.SQL = strSql
	
			qd.Parameters(0).Value = IDgroup
			qd.Parameters(1).Value = Pp.Month
			qd.Parameters(2).Value = Pp.Year
			qd.Parameters(3).Value = rs.fields("ID")
			qd.Parameters(4).Value = arrFot(j,1) ' ID FOT
	
			Set rs2 = qd.OpenRecordSet

			If Not rs2.eof Then
				If arrFot(i,4) > rs2.fields("TOTSUM") Then
					Sh.Cell(i+pos,4+j).Value = rs2.fields("TOTSUM")*arrFot(j,3)/100
				Else
					Sh.Cell(i+pos,4+j).Value = rs2.fields("TOTSUM")
				End If
			Else
				Sh.Cell(i+pos,4+j).Value = 0
			End If
		Next

		i = i + 1
		rs.moveNext
	Wend

	Sh.Rows = i + pos

	' Get Sum Columns
	For i =EE.ID) Left Join AGENTS On AGENTS.AG_ID = P_JOURNAL.AG_ID" +_
				" WHERE MTD_TYPE = 101 And P0 = [OrgID] And PP_MONTH=[mm] And PP_YEAR=[yy] And P_JOURNAL.AG_ID = [AgID] And MTD_ID_RIGHT=[FotID] " +_
				sqlAdd1 + sqlAdd8 +_
				" GROUP BY P_JOURNAL.AG_ID, MTD_ID_RIGHT, PP_MONTH, PP_YEAR, AG_DISAB, AG_AGRMNT ;" 

			Set db = workArea.DAODataBase
			Set qd = db.CreateQueryDef("")
			qd.SQL = strSql
	
			qd.Parameters(0).Value = IDgroup
			qd.Parameters(1).Value = Pp.Month
			qd.Parameters(2).Value = Pp.Year
			qd.Parameters(3).Value = rs.fields("ID")
			qd.Parameters(4).Value = arrFot(j,1) ' ID FOT
	
			Set rs2 = qd.OpenRecordSet

			If Not rs2.eof Then
				If arrFot(i,4) > rs2.fields("TOTSUM") Then
					Sh.Cell(i+pos,4+j).Value = rs2.fields("TOTSUM")*arrFot(j,3)/100
				Else
					Sh.Cell(i+pos,4+j).Value = rs2.fields("TOTSUM")
				End If
			Else
				Sh.Cell(i+pos,4+j).Value = 0
			End If
		Next

		i = i + 1
		rs.moveNext
	Wend

	Sh.Rows = Sh.Rows + 1

	' Get Sum Columns
	For