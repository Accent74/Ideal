аЯрЁБс                >  ўџ	               "          $      ўџџџ    #   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ_
		"SELECT SUM(JP_SUM) AS PSUM FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			"And AG_ID = " & agID & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTD_ID = " & mtdID )
	If Not rs.eof	Then _
		rez = rs.fields("PSUM")
	GetSumByMtd = rez
End Function

'=================
Function GetSumByCode(agID, codeID)
	Dim rez
	rez = 0
	Dim rs
	' Ьюцхђ сћђќ эхёъюыќъю ёђ№юъ/я№ютюфюъ
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT SUM(JP_SUM) AS PSUM FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			"And AG_ID = " & agID & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & codeID )
	If Not rs.eof	Then _
		rez = rs.fields("PSUM")
	GetSumByCode = rez
End Function

'=================
Function GetHoursByCode(agID, codeID)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT SUM(JP_HOURS) as REZ FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			"And AG_ID = " & agID & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & codeID )
	If Not rs.eof	Then _
		rez = rs.fields("REZ")		
	GetHoursByCode = Round2(rez / 10,2)
End Function
'=================
Function GetDaysByCode(agID, codeID)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT SUM(JP_DAYS) as REZ FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			"And AG_ID = " & agID & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & codeID )
	If Not rs.eof	Then _
		rez = rs.fields("REZ")		
	GetDaysByCode = Round2(rez / 10,2)
End Function

'=================
Function GetGHoursByCode(agID, codeID)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT SUM(JP_GHOURS) as REZ FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			"And AG_ID = " & agID & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & codeID )
	If Not rs.eof	Then _
		rez = rs.fields("REZ")		
	GetGHoursByCode = Round2(rez / 10,2)
End Function

'=================
Function GetDaysByMtd(agID, mtdID)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT SUM(JP_DAYS) as REZ FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			"And AG_ID = " & agID & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTD_ID = " & mtdID )
	If Not rs.eof	Then _
		rez = rs.fields("REZ")		
	If rez > 0 Then
		rez = Round2(rez/10)
		rez = CStr(rez)
	End If
	GetDaysByMtd =  rez
End Function

Function GetAgList
	Dim DB, QD

	Set DB = Workarea.DaoDatabase

	Set QD = DB.CreateQueryDef("")
	QD.SQL = "select * from Rep_total"
	QD.Parameters(0) = Workarea.Period.ID
	QD.Parameters(1) = 0
	QD.Parameters(2) = nved

	Set RS = QD.OpenRecordset
		GetAgList = 0
	If Not rs.eof	Then _
		GetAgList = RS.RecordCount
End Function

<     ш                Ь Arial D:\STDs.!!!\Zap\Шфхры\Шфхры       џџџџ                 џџ 
 CSheetPageSheet1Ышёђ 1         61    7                                        к   %   4      +      4      %   +   %   +   #           !   '   $   9   %   /   :   :   %       :   -   -   @              .          џџ  CRow џџ  CCell6                                                                    D K O D _ 7                               S K O D _ 7                               D A Y M T D _ 8 9                               S M T D _ 8 9                               D A Y M T D _ 1 4 3                               S M T D _ 1 4 3                               S K O D _ 1 5 7                     %         D A Y M T D _ 6 7                     %         S M T D _ 6 7                     %         D A Y M T D _ 6 5                     %         S M T D _ 6 5                        џџ      D A Y M T D _ 2 5 2                        џџ      S M T D _ 2 5 2                        џџ      D A Y M T D _ 2 5 5                        џџ      S M T D _ 2 5 5                               S M T D _ 2 5 4                               S M T D _ 2 5 8                     %         R E Z +                     %         S K O D _ 3 0 9                               S K O D _ 3 1 1                               S K O D _ 3 1 3                               S K O D _ 3 0 6                               S K O D _ 1 5 5                     %         S K O D _ 1 5 3                     %         R E Z -                        џџ      S K O D _ 1 6 2                               S K O D _ 1 1 0                               R E Z                     %                                  џџ   ~   54><>ABL  =0G8A;5=8O  70@01>B=>9  ?;0BK  !0 0 7 2   70  ?@5;L,   2 0 0 5   3.                           џџџџ                          џџџџ                          џџџџ                           5 џџ   $   (B0B=K5  A>B@C4=8:8                    5 џџ      $0<8;8O  . .                       5 џџ      $0:B                    %        '!                    %                           %                           %                             џџџџ                       %                           %                           %                           %                           %                             џџџџ                         џџџџ                         џџџџ                         џџџџ                         џџ                           џџџџ                       %     
   "                    %        #                     %                            %                              џџ                         %                            %                                      B>3>  C45@-   60=>                    % џџ     K?;0G5=>                    5 џџ                         5           2K?-   ;0B5                    % џџџџ                            5 џџ                           5 џџ                           5 џџ                        % џџ      A=.                       % џџ      B?CA:=K5                    % џџ                           џџџџ                         џџџџ                       % џџ      =4                    % џџ      >;L=8G=K5                    % џџ                         % џџ                         % џџ                           џџџџ                         џџџџ                         џџџџ                         џџџџ                       % џџ      5:@5B                    5 џџ      KE.   ?>A>1                    %                            џџ        !                     џџ        !( 1)                      џџ        !( )                     % џџ        $                     џџ      ;8<-   BK                    % џџ      @>G                                                 џџ                       5 џџ                        5 џџ                           % џџџџ                            5 џџ                           5 џџ                           5 џџ                        % џџ                        % џџ   
   0AB.                     % џџ                         % џџ      1C4                      џџџџ                       % џџ                        % џџ      @54?@.                     % џџ                            % џџ      ?>  "                    % џџ                              џџ      >  15@5<                      џџ                           џџ      >  B@02<5                      џџ                           џџ                         џџџџ                      %                           % џџ      3 , 6 %                     % џџ      2 %                      џџ      2 , 6 %                     % џџ      1 5 %                     % џџ                        % џџ                        % џџ                        % џџ  
   20=A                    5 џџ      7/ ?                    5 џџ                           % џџџџ                            5 џџ                           5 џџ   4   10H8=0  0B0;O  5==04VW2=0                    6 џџ        4@                    & џџ                        & џџ                         & џџ                         & џџ                            џџ                          & џџ                        & џџ                         & џџ                         & џџ                         & џџ                           џџ                            џџ                            џџ                            џџ                            џџ                         џџџџ                      &                           % џџ                        % џџ                        % џџ      2 %                     ! џџ                        & џџ                        & џџ                        & џџ                          џџ                       6 џџ                        6 џџ                           % џџџџ                                          <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                               2   2       <     ш              Ь Arial  срчр фрээћѕ                  <     ш               Ь Arial џџџџ                       <     ш   Ш           Ь Arial                              <     ш            џџџ Ь Arial ш \БI    К$ АБI     <     ш            РРР Ь Arial в ЌВI    b!  ГIЭЭЭЭЭЭЭЭ <     ш   №           Ь Arial рџђрђќ ярэхыќ шэёђ№ѓьхэђют  <     ш            ццњ Ь Arial 0010011:031008C9:00000000   	    џџџџ                                џџџџ                               џџџџ                                џџџџ                              џџџџ                                                                    џџџџ                               џџџџ                             џџџ                                   5  A>B@C4=8:8                    5 џџ      $0<8;8O  . .          R o o t   E n t r y                                               џџџџџџџџ                               р@ђoѕЫ         C o n t e n t s                                                  џџџџ   џџџџ                                       q;       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
                                    ўџџџўџџџ§џџџўџџџўџџџ                   !       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        арёїхђэрџ тхфюьюёђќ2011.ash                                                         413 @   `Јy`р   @   №ЃшoѕЫ@   `јЛ
ІФ      Ръіхэђ 7.0                                                                                                                                                                           џЬ' У№рєшъ - PROP_LONG - 3
' Фюыцэюёђќ - PROP_STRING - 8
' Юъырф - PROP_CY - 4

' зрёћ юђ№рсюђрээћх ёїшђрўђёџ яю ЪЮФг Юёэютэрџ Чр№яырђр

Option Explicit

Dim Sh, RS
Dim agarr, nved
'=================
Sub ShtBook_OnLoad
	Dim i
	Dim pos
	Dim rez
	Dim ri
	Dim ag
	Dim j

	Set Sh = ActiveSheet
	pos = 6 ' Ъюышїхёђтю ёђ№юъ фыџ јряъш

	Sh.Rows = pos
	NVed = InputBox("Ттхфшђх эюьх№ №рёїхђэющ тхфюьюёђш", "Й №рёїхђэющ тхфюьюёђш")
	If NVEd = "" Then 
		ShtBook.Close
		Exit Sub
	End If

Dim R
	R = GetAgList
	
	Sh.Cell(3,1).Value = "иђрђэћх ёюђ№ѓфэшъш"
	Sh.Cell(2,1).Value =  "Тхфюьюёђќ эрїшёыхэшџ чр№рсюђэющ яырђћ Й" + CStr(nved) +  " чр "+ MonthName(workarea.period.month)&", "&workarea.period.Year&" у. "

	Sh.Rows = R + pos
	rs.movefirst
	i=0
	While Not rs.eof 
	 i = i + 1
		Set ag = Workarea.Agent(RS.Fields(0).Value)
		Sh.Cell(i+pos,1).Value = i					' Яю№џфюъютћщ эюьх№
		Sh.Cell(i+pos,2).Value = RS.Fields(1).Value		' Шьџ ёюђ№ѓфэшър

		' Я№ютх№шь јряъѓ ш чряюыэшь яю єю№ьѓырь ёђюысіћ
		For j = 1 To sh.columns
			If Len(Sh.Cell(1,j).Value) > 0 Then
				Sh.Cell(i+pos,j).Value = DoCalc(ag,Sh.Cell(1,j).Value)
			End If
		Next

		rs.movenext
	Wend 

	sh.Range(1,pos,sh.columns-1,pos+r).SetBorder acBrdGrid	
	sh.Range(5,pos+1,sh.columns-1,pos+r).Alignment acRight	

	Sh.Rows = Sh.Rows + 1

	For i = 4 To Sh.Columns
		rez = 0
		For j = pos To Sh.Rows - 1
			If IsNumeric(Sh.Cell(j,i).Value) Then _
				rez = rez + CDbl(Sh.Cell(j,i).Value)
		Next
		Sh.Cell(Sh.Rows,i).Value = rez
	Next

	Sh.Cell(Sh.Rows,2).Value = "Шђюую: "
	sh.Range(8,pos,sh.columns-1,pos+r).SetBorder acBrdGrid	
	sh.Range(7,pos+1,sh.columns-1,pos+r).Alignment acRight	
	sh.Range(3,pos,sh.columns-1,pos+r+1).AutoFit acColumns
	sh.Range(2,Sh.Rows,sh.columns,Sh.Rows).Font.Bold = True

End Sub

'=================
Function DoCalc(who,what)
	Dim rez
	Dim func, arg, arg2
	Dim repAB
	rez = ""
	func = Token(what,1,"_")
	arg = Token(what,2,"_")
	arg2 = Token(what,3,"_")
	Set repAB = Workarea.Createreport("PRepAgentBill")
	repAB.AgID = who.ID
	repAB.Build
	Select Case func
		Case "SMTD"
			rez = GetSumByMtd(who.ID,arg)
		Case "SKOD"
			rez = GetSumByCode(who.ID,arg)
		Case "REZ+"
'			rez = RS.Fields(2).Value
			rez = repAB.TotalIN
		Case "REZ-"
'			rez = RS.Fields(3).Value
			rez = repAB.TotalWD - Checknull(GetSumByCode(who.ID,162),0) - Checknull(GetSumByCode(who.ID,110),0)
		Case "REZ"
'			rez = RS.Fields(4).Value
			rez = repAB.TotalIN - repAB.TotalWD
		Case "PRM"
			rez = GetInfo(who.ID,GetPropID(arg),arg2)
		Case "GHKOD"
			rez = GetGHoursByCode(who.ID,arg)
		Case "HKOD"
			rez = GetHoursByCode(who.ID,arg)
		Case "DKOD"
			rez = GetDaysByCode(who.ID,arg)
		Case "DAYMTD"
			rez = GetDaysByMtd(who.ID,arg)
		Case "SUMM"
			rez = Checknull(GetSumByMtd(who.ID,arg),0) + Checknull(GetSumByMtd(who.ID,arg2),0)
		Case "SUMK"
			rez = Checknull(GetSumByCode(who.ID,arg),0) + Checknull(GetSumByCode(who.ID,arg2),0)
		Case Else
			rez = "?"
	End Select
	DoCalc = rez 
End Function

'=================
Function GetInfo(agID,propID,tf)
	Dim rez
	rez = 0
	
	Dim db,qd,rs
	Set db = WorkArea.DAODataBase
	Set qd = db.CreateQueryDef("")
	qd.SQL = "SELECT * FROM AgProp_Values WHERE PROP_ID = "&propID
	QD.Parameters(0).Value = agID
	QD.Parameters(1).Value = Now
	Set rs = qd.OpenRecordset
	If Not rs.eof Then
		Select Case CInt(tf)
			Case 8
				rez = rs.fields("Prop_String")	
			Case 4
				rez = rs.fields("Prop_Cy")	
			Case 3
				rez = rs.fields("Prop_Long")	
		End Select
	End If
	GetInfo = rez
End Function

'=================
Function GetPropID(propName)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT PROP_ID FROM AG_PROP_NAMES WHERE PROP_NAME = '"&propName&"'")
	If Not rs.eof	Then _
		rez = rs.fields("PROP_ID")	
	GetPropID = rez
End Function

'=================
Function GetSumByMtd(agID, mtdID)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT SUM(JP_SUM) AS PSUM FROM P_JOURNAR o o t   E n t r y                                               џџџџџџџџ                                С  Ь%         C o n t e n t s                                                  џџџџ   џџџџ                                    &   ;       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
                                    ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ'   (   )   *   +   ,   -       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        арёїхђэрџ тхфюьюёђќ2011.ash                                                         414 @   рd]с   @   Р8А  Ь@   `јЛ
ІФ      Ръіхэђ 7.0                                                                                                                                                                           џњ' У№рєшъ - PROP_LONG - 3
' Фюыцэюёђќ - PROP_STRING - 8
' Юъырф - PROP_CY - 4

' зрёћ юђ№рсюђрээћх ёїшђрўђёџ яю ЪЮФг Юёэютэрџ Чр№яырђр

Option Explicit

Dim Sh, RS
Dim agarr, nved
'=================
Sub ShtBook_OnLoad
	Dim i
	Dim pos
	Dim rez
	Dim ri
	Dim ag
	Dim j

	Set Sh = ActiveSheet
	pos = 6 ' Ъюышїхёђтю ёђ№юъ фыџ јряъш

	Sh.Rows = pos
	NVed = InputBox("Ттхфшђх эюьх№ №рёїхђэющ тхфюьюёђш", "Й №рёїхђэющ тхфюьюёђш")
	If NVEd = "" Then 
		ShtBook.Close
		Exit Sub
	End If

Dim R
	R = GetAgList
	
	Sh.Cell(3,1).Value = "иђрђэћх ёюђ№ѓфэшъш"
	Sh.Cell(2,1).Value =  "Тхфюьюёђќ эрїшёыхэшџ чр№рсюђэющ яырђћ Й" + CStr(nved) +  " чр "+ MonthName(workarea.period.month)&", "&workarea.period.Year&" у. "

	Sh.Rows = R + pos
	rs.movefirst
	i=0
	While Not rs.eof 
	 i = i + 1
		Set ag = Workarea.Agent(RS.Fields(0).Value)
		Sh.Cell(i+pos,1).Value = i					' Яю№џфюъютћщ эюьх№
		Sh.Cell(i+pos,2).Value = RS.Fields(1).Value		' Шьџ ёюђ№ѓфэшър

		' Я№ютх№шь јряъѓ ш чряюыэшь яю єю№ьѓырь ёђюысіћ
		For j = 1 To sh.columns
			If Len(Sh.Cell(1,j).Value) > 0 Then
				Sh.Cell(i+pos,j).Value = DoCalc(ag,Sh.Cell(1,j).Value)
			End If
		Next

		rs.movenext
	Wend 

	sh.Range(1,pos,sh.columns-1,pos+r).SetBorder acBrdGrid	
	sh.Range(5,pos+1,sh.columns-1,pos+r).Alignment acRight	

	Sh.Rows = Sh.Rows + 1

	For i = 4 To Sh.Columns
		rez = 0
		For j = pos To Sh.Rows - 1
			If IsNumeric(Sh.Cell(j,i).Value) Then _
				rez = rez + CDbl(Sh.Cell(j,i).Value)
		Next
		Sh.Cell(Sh.Rows,i).Value = rez
	Next

	Sh.Cell(Sh.Rows,2).Value = "Шђюую: "
	sh.Range(8,pos,sh.columns-1,pos+r).SetBorder acBrdGrid	
	sh.Range(7,pos+1,sh.columns-1,pos+r).Alignment acRight	
	sh.Range(3,pos,sh.columns-1,pos+r+1).AutoFit acColumns
	sh.Range(2,Sh.Rows,sh.columns,Sh.Rows).Font.Bold = True

End Sub

'=================
Function DoCalc(who,what)
	Dim rez
	Dim func, arg, arg2
	Dim repAB
	rez = ""
	func = Token(what,1,"_")
	arg = Token(what,2,"_")
	arg2 = Token(what,3,"_")
	Set repAB = Workarea.Createreport("PRepAgentBill")
	repAB.AgID = who.ID
	repAB.Build
	Select Case func
		Case "SMTD"
			rez = GetSumByMtd(who.ID,arg)
		Case "SKOD"
			rez = GetSumByCode(who.ID,arg)
		Case "REZ+"
'			rez = RS.Fields(2).Value
			rez = repAB.TotalIN
		Case "REZ-"
'			rez = RS.Fields(3).Value
			rez = repAB.TotalWD - Checknull(GetSumByCode(who.ID,Workarea.GetCodeID("3020")),0) - Checknull(GetSumByCode(who.ID,Workarea.GetCodeID("3010")),0)
		Case "REZ"
'			rez = RS.Fields(4).Value
			rez = repAB.TotalIN - repAB.TotalWD
		Case "PRM"
			rez = GetInfo(who.ID,GetPropID(arg),arg2)
		Case "GHKOD"
			rez = GetGHoursByCode(who.ID,arg)
		Case "HKOD"
			rez = GetHoursByCode(who.ID,arg)
		Case "DKOD"
			rez = GetDaysByCode(who.ID,arg)
		Case "DAYMTD"
			rez = GetDaysByMtd(who.ID,arg)
		Case "SUMM"
			rez = Checknull(GetSumByMtd(who.ID,arg),0) + Checknull(GetSumByMtd(who.ID,arg2),0)
		Case "SUMK"
			rez = Checknull(GetSumByCode(who.ID,arg),0) + Checknull(GetSumByCode(who.ID,arg2),0)
		Case Else
			rez = "?"
	End Select
	DoCalc = rez 
End Function

'=================
Function GetInfo(agID,propID,tf)
	Dim rez
	rez = 0
	
	Dim db,qd,rs
	Set db = WorkArea.DAODataBase
	Set qd = db.CreateQueryDef("")
	qd.SQL = "SELECT * FROM AgProp_Values WHERE PROP_ID = "&propID
	QD.Parameters(0).Value = agID
	QD.Parameters(1).Value = Now
	Set rs = qd.OpenRecordset
	If Not rs.eof Then
		Select Case CInt(tf)
			Case 8
				rez = rs.fields("Prop_String")	
			Case 4
				rez = rs.fields("Prop_Cy")	
			Case 3
				rez = rs.fields("Prop_Long")	
		End Select
	End If
	GetInfo = rez
End Function

'=================
Function GetPropID(propName)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT PROP_ID FROM AG_PROP_NAMES WHERE PROP_NAME = '"&propName&"'")
	If Not rs.eof	Then _
		rez = rs.fields("PROP_ID")	
	GetPropID = rez
End Function

'=================
Function GetSumByMtd(agID, mtdID)
	Dim rez
	rez = 0
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( 