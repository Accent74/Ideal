аЯрЁБс                >  ўџ	                         	      ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ
	For i = 0 To UBound(PrmValues)
		CreateParameter Cmd, PrmValues(i), PrmTypes(i)
	Next

	Cmd.CommandType = adCmdText 
	Set ExecuteQuery = Cmd.Execute 
End Function

'бючфрхь тѕюфџљшщ яр№рьхђ№ ёю чэрїхэшхь ParamValue ш ђшяюь ParamType ---------------------------
Sub CreateParameter(Cmd, ParamValue, ParamType)
	Dim Prm

	Set Prm = Cmd.CreateParameter(, ParamType, adParamInput)
	Prm.Value = ParamValue
	Cmd.Parameters.Append Prm
End Sub
<     ш                Ь Arial D:\STDs.!!!\Zap\Шфхры\Шфхры       џџџџ                 џџ 
 CSheetPageSheet1Тхфюьюёђќ эрїшёыхэшщ         61    7                                     k   3                    џџ  CRow џџ  CCell6     џџ   (   54><>ABL  =0G8A;5=89                        џџ      /=20@L  2 0 0 6   3                     % 	        >4@0745;5=85                    % џџ    
   !C<<0                     % 	        !>B@C4=8:                    ! џџ       ;L3>BK                             B>3>:                     "                        d   d   d   d    <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                 2   2   
    <     ш    Ш           Ь Arial аDb                   
    <     ш   №           Ь Arial Sub ShtBook_OnLoadEnd Su  <     ш    №           Ь Arial шђх F1 ыхээюую §ыхьхэђр 
.av <     ш               Ь Arial шђх F1  (F7)    1      шЭe <     ш           џџр Ь Arial џџ               №ѓђъш    <     ш           ((( Ь Arial џџ               htRange) <     ш           џџџ Ь Arial џџ                   ing  <     ш            ((( Ь Arial џџ                   а#g <     ш   Ш           Ь Arial                      шЭe	 <     ш    Д           Ь Arial ющёђт ћ (F7) ёюёђюџэшџ фыџ ю    џџџ                               ШШШ                               ШШШ                                  ммм                                  ммм                              џџр                                  џџџџ                             џџр                              џџџџ                              џџџџ                              	 џџџџ                             
 рџџ                               џџџ                                рџџ                                рџџ                               рџџ                              џџџ                                џџџ                                 рџџ                                 џџџ                                џџџ                               џџџџ                               џџџџ                                џџџџ                               РРР                                  ШШШ                               РРР                               ШШШ                               РРР                              џџџ                                  џџџ                               #45@60=>                      џџ       #45@60=>                      џџ       #45@60=>                      џџ    
   B>3>                      џџ       2KR o o t   E n t r y                                               џџџџџџџџ                               №НEдФЧ
         C o n t e n t s                                                  џџџџ   џџџџ                                       &M       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                  ўџџџўџџџ§џџџўџџџўџџџ                           )   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ*   +   ,   -   .   /   0   1   2   3   4   5   6   7   8   9   :   ;   <   =   >       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Тхфюьюёђќ эрїшёыхэшщ.ash                                                            260 @   @Eуњс   @   ѕ4дФЧ@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                           џУAOption Explicit

Const adCmdText = 1
Const adParamInput = 1

Const StartNumberOfRows = 5
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "дЮв"

Dim cRow
Dim Sht

Dim DicCodes
Dim TotalsPos

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Set Sht = ShtBook.Sheet1
	Set DicCodes = CreateObject("Scripting.Dictionary")

	Sht.DisplayGrid = False
	Sht.DisplayHeadings = False

	BuildReport
End Sub

'-----------------------------------------------------------------------------------------------
Sub BuildReport
	OutputRepPeriod
	BuildMainReport
	BuildSecondaryReport
End Sub

'-----------------------------------------------------------------------------------------------
Sub OutputRepPeriod
	Dim FirstDay

	FirstDay = WorkArea.Period.RepFirstDay

	Sht.Cell(2, 1).Value =  MonthName(Month(FirstDay)) & " " & Year(FirstDay) & " у"
End Sub

'-----------------------------------------------------------------------------------------------
Sub BuildMainReport
	Sht.Rows = StartNumberOfRows
	Sht.Columns = StartNumberOfColumns

	Sht.Rows = StartNumberOfRows
	cRow = StartNumberOfRows - 1

	DicCodes.RemoveAll
	ReDim TotalsPos(0)

	TotalsPos(0) = StartNumberOfColumns

	SetColumns
	SetTotalsRow

	FillData
	PrepareReportForShowing
End Sub

'-----------------------------------------------------------------------------------------------
Sub SetColumns
	AddCodes 101, "Эрїшёыхэю"
	AddCodes 102, "гфх№црэю"

	Sht.InsertColumn

	With Sht.Range(Sht.Columns, cRow-1, Sht.Columns, cRow)
		.MergeCells = True
		.Value = "Ъ тћфрїх"
		.SetBorder acBrdGrid
	End With

	Sht.Range(StartNumberOfColumns+1, cRow-1, Sht.Columns, cRow).BackColor = RGB(200, 200, 200)
End Sub

'-----------------------------------------------------------------------------------------------
Sub AddCodes(CodeType, NameOfType)
	Dim i
	Dim Codes
	Dim NumberOfColumns

	Set Codes = GetCodes(CodeType)

	NumberOfColumns = Sht.Columns
	Sht.Columns = NumberOfColumns + Codes.RecordCount + 1

	For i = NumberOfColumns + 1 To Sht.Columns - 1
		Sht.Cell(cRow, i).Value = Codes.Fields(0).Value

		DicCodes.Add Codes.Fields(1).Value, i
		Codes.MoveNext
	Next

	Sht.Range(NumberOfColumns+1, cRow, Sht.Columns, cRow).Alignment = acCenter
	Sht.Range(NumberOfColumns+1, cRow-1, Sht.Columns-1, cRow-1).Value = NameOfType

	ReDim Preserve TotalsPos(UBound(TotalsPos)+1)
	TotalsPos(UBound(TotalsPos)) = Sht.Columns

	Sht.Range(TotalsPos(UBound(TotalsPos)-1)+1, cRow-1, Sht.Columns-1, cRow).SetBorder acBrdGrid

	Sht.Cell(cRow-1, Sht.Columns).Value = "Шђюую"
	Sht.Cell(cRow, Sht.Columns).Value = LCase(NameOfType)
	Sht.Range(Sht.Columns, cRow-1, Sht.Columns, cRow).SetBorder acBrdOutline
End Sub

'-----------------------------------------------------------------------------------------------
Function GetCodes(CodeType)
	Dim SqlQr

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS MtdType Long; " & _
				       "Select MTD_CODE, MTD_ID From MTD_CODES Where MTD_TYPE = [MtdType] Order By MTD_CODE"

				.Parameters(0).Value = CodeType

				Set GetCodes = .OpenRecordSet
			End With
		Case "OLEDB"
			SqlQr = "Select MTD_CODE, MTD_ID From MTD_CODES Where MTD_TYPE = ? Order By MTD_CODE"

			Set GetCodes = ExecuteQuery(SqlQr, Array(CodeType), Array(vbLong))
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Sub SetTotalsRow
	cRow = cRow + 1

	With Sht.Range(StartNumberOfColumns+1, cRow, Sht.Columns, cRow)
		.Font.Bold = True
		.Alignment = acRight
		.BackColor = Sht.Cell(cRow, 1).BackColor
		.SetBorder acBrdGrid
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Sub FillData
	Dim RecSet
	Dim GroupID
	Dim AgID
	Dim PeriodID
	Dim PrivilegeID

	PeriodID = GetRepPeriodID(WorkArea.Period)
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)

	Set RecSet = GetData(PeriodID)

	GroupID = 0
	AgID = 0

	With RecSet
		Do While Not .EOF
			If GroupID <> .Fields(0).Value Then
				cRow = Sht.InsertRow2

				Sht.Cell(cRow, 1).CellParam = 1
				Sht.Range(1, cRow, Sht.Columns, cRow).Font.Bold = True

				Sht.Cell(cRow, 1).Value = .Fields(2).Value
				GroupID = .Fields(0).Value
			End If

			If AgID <> .Fields(1).Value Then
				cRow = Sht.InsertRow2

				Sht.Cell(cRow, 1).Value = .Fields(3).Value
				Sht.Cell(cRow, 2).Value = GetPrivilege(PrivilegeID, .Fields(1).Value, PeriodID)
				AgID = .Fields(1).Value
			End If

			Sht.Cell(cRow, DicCodes(.Fields(4).Value)).Value = .Fields(5).Value

			.MoveNext
		Loop
	End With

	CalcTotals
End Sub

'-----------------------------------------------------------------------------------------------
Function GetRepPeriodID(Period)
	Dim SqlQr
	Dim PDay

	PDay = Period.RepFirstDay

	Select Case App.AppType
		Case "DAO"
			With WorkR o o t   E n t r y                                               џџџџџџџџ                               №>ЗФЧ         C o n t e n t s                                                  џџџџ   џџџџ                                       M       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                  ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ                      !   "   #   $   %   &   '   (   ?   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ@   A   B   C   D   E   F   G   H   I   J   K   L   M   N       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Тхфюьюёђќ эрїшёыхэшщ.ash                                                            259 @   P9онс   @    щ/ЗФЧ@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                           џУAOption Explicit

Const adCmdText = 1
Const adParamInput = 1

Const StartNumberOfRows = 5
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "дЮв"

Dim cRow
Dim Sht

Dim DicCodes
Dim TotalsPos

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Set Sht = ShtBook.Sheet1
	Set DicCodes = CreateObject("Scripting.Dictionary")

	Sht.DisplayGrid = False
	Sht.DisplayHeadings = False

	BuildReport
End Sub

'-----------------------------------------------------------------------------------------------
Sub BuildReport
	OutputRepPeriod
	BuildMainReport
	BuildSecondaryReport
End Sub

'-----------------------------------------------------------------------------------------------
Sub OutputRepPeriod
	Dim FirstDay

	FirstDay = WorkArea.Period.RepFirstDay

	Sht.Cell(2, 1).Value =  MonthName(Month(FirstDay)) & " " & Year(FirstDay) & " у"
End Sub

'-----------------------------------------------------------------------------------------------
Sub BuildMainReport
	Sht.Rows = StartNumberOfRows
	Sht.Columns = StartNumberOfColumns

	Sht.Rows = StartNumberOfRows
	cRow = StartNumberOfRows - 1

	DicCodes.RemoveAll
	ReDim TotalsPos(0)

	TotalsPos(0) = StartNumberOfColumns

	SetColumns
	SetTotalsRow

	FillData
	PrepareReportForShowing
End Sub

'-----------------------------------------------------------------------------------------------
Sub SetColumns
	AddCodes 101, "Эрїшёыхэю"
	AddCodes 102, "гфх№црэю"

	Sht.InsertColumn

	With Sht.Range(Sht.Columns, cRow-1, Sht.Columns, cRow)
		.MergeCells = True
		.Value = "Ъ тћфрїх"
		.SetBorder acBrdGrid
	End With

	Sht.Range(StartNumberOfColumns+1, cRow-1, Sht.Columns, cRow).BackColor = RGB(200, 200, 200)
End Sub

'-----------------------------------------------------------------------------------------------
Sub AddCodes(CodeType, NameOfType)
	Dim i
	Dim Codes
	Dim NumberOfColumns

	Set Codes = GetCodes(CodeType)

	NumberOfColumns = Sht.Columns
	Sht.Columns = NumberOfColumns + Codes.RecordCount + 1

	For i = NumberOfColumns + 1 To Sht.Columns - 1
		Sht.Cell(cRow, i).Value = Codes.Fields(0).Value

		DicCodes.Add Codes.Fields(1).Value, i
		Codes.MoveNext
	Next

	Sht.Range(NumberOfColumns+1, cRow, Sht.Columns, cRow).Alignment = acCenter
	Sht.Range(NumberOfColumns+1, cRow-1, Sht.Columns-1, cRow-1).Value = NameOfType

	ReDim Preserve TotalsPos(UBound(TotalsPos)+1)
	TotalsPos(UBound(TotalsPos)) = Sht.Columns

	Sht.Range(TotalsPos(UBound(TotalsPos)-1)+1, cRow-1, Sht.Columns-1, cRow).SetBorder acBrdGrid

	Sht.Cell(cRow-1, Sht.Columns).Value = "Шђюую"
	Sht.Cell(cRow, Sht.Columns).Value = LCase(NameOfType)
	Sht.Range(Sht.Columns, cRow-1, Sht.Columns, cRow).SetBorder acBrdOutline
End Sub

'-----------------------------------------------------------------------------------------------
Function GetCodes(CodeType)
	Dim SqlQr

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS MtdType Long; " & _
				       "Select MTD_CODE, MTD_ID From MTD_CODES Where MTD_TYPE = [MtdType] Order By MTD_CODE"

				.Parameters(0).Value = CodeType

				Set GetCodes = .OpenRecordSet
			End With
		Case "OLEDB"
			SqlQr = "Select MTD_CODE, MTD_ID From MTD_CODES Where MTD_TYPE = ? Order By MTD_CODE"

			Set GetCodes = ExecuteQuery(SqlQr, Array(CodeType), Array(vbLong))
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Sub SetTotalsRow
	cRow = cRow + 1

	With Sht.Range(StartNumberOfColumns+1, cRow, Sht.Columns, cRow)
		.Font.Bold = True
		.Alignment = acRight
		.BackColor = Sht.Cell(cRow, 1).BackColor
		.SetBorder acBrdGrid
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Sub FillData
	Dim RecSet
	Dim GroupID
	Dim AgID
	Dim PeriodID
	Dim PrivilegeID

	PeriodID = GetRepPeriodID(WorkArea.Period)
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)

	Set RecSet = GetData(PeriodID)

	GroupID = 0
	AgID = 0

	With RecSet
		Do While Not .EOF
			If GroupID <> .Fields(0).Value Then
				cRow = Sht.InsertRow2

				Sht.Cell(cRow, 1).CellParam = 1
				Sht.Range(1, cRow, Sht.Columns, cRow).Font.Bold = True

				Sht.Cell(cRow, 1).Value = .Fields(2).Value
				GroupID = .Fields(0).Value
			End If

			If AgID <> .Fields(1).Value Then
				cRow = Sht.InsertRow2

				Sht.Cell(cRow, 1).Value = .Fields(3).Value
				Sht.Cell(cRow, 2).Value = GetPrivilege(PrivilegeID, .Fields(1).Value, PeriodID)
				AgID = .Fields(1).Value
			End If

			Sht.Cell(cRow, DicCodes(.Fields(4).Value)).Value = .Fields(5).Value

			.MoveNext
		Loop
	End With

	CalcTotals
End Sub

'-----------------------------------------------------------------------------------------------
Function GetRepPeriodID(Period)
	Dim SqlQr
	Dim PDay

	PDay = Period.RepFirstDay

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS PMonth Long, PYear Long; " & _
				       "Select PP_ID From P_PERIOD Where PP_MONTH = [PMonth] And PP_YEAR = [PYear]"

				.Parameters(0).Value = Month(PDay)
				.Parameters(1).Value = Year(PDay)

				GetRepPeriodID = .OpenRecordSet.Fields(0).Value
			End With
		Case "OLEDB"
			SqlQr = "Select PP_ID From P_PERIOD Where PP_MONTH = ? And PP_YEAR = ?"
			GetRepPeriodID = ExecuteQuery(SqlQr, Array(Month(PDay), Year(PDay)), Array(vbLong, vbLong)).Fields(0).Value
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Function GetData(PeriodID)
	Dim SqlQr

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS PeriodID Long; " & _
				       "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
				       "From MTD_CODES Inner Join " & _
				       "    (P_JOURNAL Inner Join " & _
				       "       (AGENTS Inner Join " & _
				       "      (AG_TREE Inner Join AGENTS Groups On AG_TREE.P0 = Groups.AG_ID) " & _
				       "                                        On AGENTS.AG_ID = AG_TREE.ID) " & _
				       "                                        On P_JOURNAL.AG_ID = AGENTS.AG_ID) " & _
				       "                                        On MTD_CODES.MTD_ID = P_JOURNAL.MTC_ID " & _
				       "Where JP_DONE = 1 And SHORTCUT = 0 And (MTD_TYPE = 101 Or MTD_TYPE = 102) " & _
				       "                  And (MTD_MASK < 8 Or IsNull(MTD_MASK))" & _
				       "                  And C_PERIOD = [PeriodID] And W_PERIOD = [PeriodID] " & _
				       "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
				       "Having Sum(JP_SUM) <> 0 " & _
				       "Order By Groups.AG_NAME, AGENTS.AG_NAME "

				.Parameters(0).Value = PeriodID

				Set GetData = .OpenRecordSet
			End With
		Case "OLEDB"
			SqlQr = "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
			        "From P_JOURNAL Inner Join AGENTS On P_JOURNAL.AG_ID = AGENTS.AG_ID " & _
			        "               Inner Join AG_TREE On P_JOURNAL.AG_ID = ID " & _
			        "               Inner Join MTD_CODES On MTC_ID = MTD_ID " & _
			        "               Left Join AGENTS Groups On P0 = Groups.AG_ID " & _
			        "Where JP_DONE = 1 And SHORTCUT = 0 And (MTD_TYPE = 101 Or MTD_TYPE = 102) " & _
			        "                  And IsNull(MTD_MASK, 0) < 8 " & _
			        "                  And C_PERIOD = ? And W_PERIOD = ? " & _
			        "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
			        "Having Sum(JP_SUM) <> 0 " & _
			        "Order By Groups.AG_NAME, AGENTS.AG_NAME "

			Set GetData = ExecuteQuery(SqlQr, Array(PeriodID, PeriodID), Array(vbLong, vbLong))
	End Select
End Function

'-----------------------------------------------------------------------------------------------
FunctArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS PMonth Long, PYear Long; " & _
				       "Select PP_ID From P_PERIOD Where PP_MONTH = [PMonth] And PP_YEAR = [PYear]"

				.Parameters(0).Value = Month(PDay)
				.Parameters(1).Value = Year(PDay)

				GetRepPeriodID = .OpenRecordSet.Fields(0).Value
			End With
		Case "OLEDB"
			SqlQr = "Select PP_ID From P_PERIOD Where PP_MONTH = ? And PP_YEAR = ?"
			GetRepPeriodID = ExecuteQuery(SqlQr, Array(Month(PDay), Year(PDay)), Array(vbLong, vbLong)).Fields(0).Value
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Function GetData(PeriodID)
	Dim SqlQr

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS PeriodID Long; " & _
				       "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
				       "From MTD_CODES Inner Join " & _
				       "    (P_JOURNAL Inner Join " & _
				       "       (AGENTS Inner Join " & _
				       "      (AG_TREE Inner Join AGENTS Groups On AG_TREE.P0 = Groups.AG_ID) " & _
				       "                                        On AGENTS.AG_ID = AG_TREE.ID) " & _
				       "                                        On P_JOURNAL.AG_ID = AGENTS.AG_ID) " & _
				       "                                        On MTD_CODES.MTD_ID = P_JOURNAL.MTC_ID " & _
				       "Where JP_DONE = 1 And SHORTCUT = 0 And (MTD_TYPE = 101 Or MTD_TYPE = 102) " & _
				       "                  And (MTD_MASK < 8 Or IsNull(MTD_MASK))" & _
				       "                  And C_PERIOD = [PeriodID] And W_PERIOD = [PeriodID] " & _
				       "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
				       "Having Sum(JP_SUM) <> 0 " & _
				       "Order By Groups.AG_NAME, AGENTS.AG_NAME "

				.Parameters(0).Value = PeriodID

				Set GetData = .OpenRecordSet
			End With
		Case "OLEDB"
			SqlQr = "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
			        "From P_JOURNAL Inner Join AGENTS On P_JOURNAL.AG_ID = AGENTS.AG_ID " & _
			        "               Inner Join AG_TREE On P_JOURNAL.AG_ID = ID " & _
			        "               Inner Join MTD_CODES On MTC_ID = MTD_ID " & _
			        "               Left Join AGENTS Groups On P0 = Groups.AG_ID " & _
			        "Where JP_DONE = 1 And SHORTCUT = 0 And (MTD_TYPE = 101 Or MTD_TYPE = 102) " & _
			        "                  And IsNull(MTD_MASK, 0) < 8 " & _
			        "                  And C_PERIOD = ? And W_PERIOD = ? " & _
			        "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
			        "Having Sum(JP_SUM) <> 0 " & _
			        "Order By Groups.AG_NAME, AGENTS.AG_NAME "

			Set GetData = ExecuteQuery(SqlQr, Array(PeriodID, PeriodID), Array(vbLong, vbLong))
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Function GetPrivilege(MtcID, AgID, PeriodID)
	Dim RecSet
	Dim SqlQr

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS MtcID Long, AgID Long, PeriodID Long; " & _
				       "Select JP_S1 From P_JOURNAL " & _
				       "Where JP_DONE = 1 And (Not IsNull(JP_S1)) And " & _
				       "MTC_ID = [MtcID] And AG_ID = [AgID] And C_PERIOD = [PeriodID] And W_PERIOD = [PeriodID]"

				.Parameters(0).Value = MtcID
				.Parameters(1).Value = AgID
				.Parameters(2).Value = PeriodID

				Set RecSet = .OpenRecordSet
			End With
		Case "OLEDB"
			Dim PrmValues
			Dim PrmTypes

			PrmValues = Array(MtcID, AgID, PeriodID, PeriodID)
			PrmTypes  = Array(vbLong, vbLong, vbLong, vbLong)

			SqlQr = "Select JP_S1 From P_JOURNAL " & _
			        "Where JP_DONE = 1 And IsNull(JP_S1, 0) <> 0 And " & _
			        "MTC_ID = ? And AG_ID = ? And C_PERIOD = ? And W_PERIOD = ? "

			Set RecSet = ExecuteQuery(SqlQr, PrmTypes, PrmTypes)
	End Select

	If RecSet.EOF Then
		GetPrivilege = 0
	Else
		GetPrivilege = RecSet.Fields(0).Value
	End If
End Function

'-----------------------------------------------------------------------------------------------
Sub CalcTotals
	Dim i, j
	ReDim GroupTotals(Sht.Columns)
	ReDim Totals(Sht.Columns)

	CalcTotalsInRows

	SetZero Totals
	SetZero GroupTotals

	For i = Sht.Rows To StartNumberOfRows + 1 Step -1
		If Sht.Cell(i, 1).CellParam = 1 Then
			For j = StartNumberOfColumns + 1 To Sht.Columns
				Sht.Cell(i, j).Value = GroupTotals(j)
				Totals(j) = Totals(j) + GroupTotals(j)
			Next

			SetZero GroupTotals
		Else
			For j = StartNumberOfColumns + 1 To Sht.Columns
				GroupTotals(j) = GroupTotals(j) + Sht.Cell(i, j).Value
			Next
		End If
	Next

	For j = StartNumberOfColumns + 1 To Sht.Columns
		Sht.Cell(StartNumberOfRows, j).Value = Totals(j)
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub CalcTotalsInRows
	Dim i, j, k
	Dim Sum

	For i = StartNumberOfRows + 1 To Sht.Rows
		For k = 1 To UBound(TotalsPos)
			Sum = 0
	
			For j = TotalsPos(k-1) + 1 To TotalsPos(k) - 1
				Sum = Sum + Sht.Cell(i, j).Value
			Next

			Sht.Cell(i, TotalsPos(k)).Value = Sum
		Next

		Sht.Cell(i, Sht.Columns).Value = Sht.Cell(i, TotalsPos(1)).Value - Sht.Cell(i, TotalsPos(2)).Value
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub SetZero(Arr)
	Dim i
	
	For i = 1 To UBound(Arr)
		Arr(i) = 0
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub PrepareReportForShowing
	Dim j

	DeleteEmptyColumns

	For j = 1 To Sht.Columns
		Sht.Range(j, StartNumberOfRows+1, j, Sht.Rows).SetBorder acBrdOutline
	Next

	Sht.Range(StartNumberOfColumns, StartNumberOfRows+1, Sht.Columns, Sht.Rows).Alignment = acRight

	MergeCells StartNumberOfRows-2, "Эрїшёыхэю"
	MergeCells StartNumberOfRows-2, "гфх№црэю"

	Sht.Range(1, StartNumberOfRows-2, Sht.Columns, Sht.Rows).Autofit
	Sht.Range(1, StartNumberOfRows+1, Sht.Columns, Sht.Rows).Stripe
End Sub

'-----------------------------------------------------------------------------------------------
Sub DeleteEmptyColumns
	Dim j

	For j = Sht.Columns To StartNumberOfColumns + 1 Step -1
		If Sht.Cell(StartNumberOfRows, j).Value = 0 Then Sht.DeleteColumn j
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub MergeCells(OnRow, CellsText)
	Dim j
	Dim FirstPos, LastPos

	FirstPos = 1
	LastPos = 0

	For j = 1 To Sht.Columns
		If Sht.Cell(OnRow, j).Value = CellsText Then
			FirstPos = j
			Exit For
		End If
	Next

	For j = Sht.Columns To 1 Step -1
		If Sht.Cell(OnRow, j).Value = CellsText Then
			LastPos = j 
			Exit For
		End If
	Next

	If LastPos >= FirstPos Then Sht.Range(FirstPos, OnRow, LastPos, OnRow).MergeCells = True
End Sub

'-----------------------------------------------------------------------------------------------
Sub BuildSecondaryReport
	Dim ColNames
	Dim ColNumber
	Dim TotalsRow

	If Sht.Rows = StartNumberOfRows Then Exit Sub

	ColNames = Array("Эрїшёыхэшх", "%", "дЮв", "бѓььр", "Фхсхђ", "Ъ№хфшђ")
	ColNumber = 2 * (UBound(ColNames) + 1)

	If Sht.Columns < ColNumber Then Sht.Columns = ColNumber

	SetHeadings ColNames, ColNumber

	TotalsRow = Sht.InsertRow2

	SetTotalsRow2 TotalsRow, ColNumber
	FillWPF TotalsRow
	EndReport TotalsRow, ColNumber
End Sub

'-----------------------------------------------------------------------------------------------
Sub SetHeadings(ColNames, ColNumber)
	Dim j

	Sht.InsertRow2
	cRow = Sht.InsertRow2

	For j = 0 To UBound(ColNames)
		Sht.Cell(cRow, 2*j+1).Value = ColNames(j)
	Next

	With Sht.Range(1, cRow, ColNumber, cRow)
		.BackColor = Sht.Cell(StartNumberOfRows-1, 1).BackColor
		.Alignment = acCenter
		.SetBorder acBrdGrid
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Sub SetTotalsRow2(Row, ColNumber)
	Sht.Cell(Row, 1).Value = "Шђюую:"

	With Sht.Range(1, Row, ColNumber, Row)
		.Font.Bold = True
		.BackColor = Sht.Cell(StartNumberOfRows, 1).BackColor
		.SetBorder acBrdGrid
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Sub FillWPF(TotalsRow)
	Dim i
	Dim WPFCodes
	Dim Sum

	Set WPFCodes = WorkArea.Code(GetCodeID(MTC_WPF_NAME)).Children

	Sum = 0

	For i = 1 To WPFCodes.Count
		cRow = Sht.InsertRow2

		With WPFCodes(i)
			Sht.Cell(cRow, 1).Value = .Name
			Sht.Cell(cRow, 3).Value = .Percent
			Sht.Cell(cRow, 5).Value = GetWPFSum(.Threshold, TotalsRow-3)
			Sht.Cell(cRow, 7).Value =  Sht.Cell(cRow, 5).Value * .Percent / 100
			Sht.Cell(cRow, 9).Value = .DbCode
			Sht.Cell(cRow, 11).Value = .CrCode
		End With

		Sum = Sum + Sht.Cell(cRow, 7).Value
	Next

	Sht.Cell(TotalsRow, 7).Value = Sum
End Sub

'-----------------------------------------------------------------------------------------------
Function GetCodeID(CodeName)
	Dim RecSet
	Dim SqlQr

	SqlQr = "Select MTD_ID From MTD_CODES Where MTD_NAME = '" & CodeName & "'"

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = SqlQr
				Set RecSet = .OpenRecordSet
			End With
		Case "OLEDB"
			Set RecSet = ExecuteQuery(SqlQr, Array(), Array())
	End Select

	If RecSet.EOF Then
		GetCodeID = 0
	Else
		GetCodeID = RecSet.Fields(0).Value
	End If
End Function

'-----------------------------------------------------------------------------------------------
Function GetWPFSum(Threshold, EndRow)
	Dim i
	Dim Col
	Dim Sum

	For Col = 1 To Sht.Columns
		If Sht.Cell(StartNumberOfRows-1, Col).Value = "эрїшёыхэю" Then Exit For
	Next

	Sum = 0

	For i = StartNumberOfRows + 1 To EndRow
		With Sht.Cell(i, Col)
			If .Font.Bold = False Then Sum = Sum + IIF(.Value < Threshold, .Value, Threshold)
		End With
	Next

	GetWPFSum = Sum
End Function

'-----------------------------------------------------------------------------------------------
Sub EndReport(TotalsRow, ColNumber)
	Dim i, j

	For j = 1 To ColNumber Step 2
		For i = TotalsRow - 1 To Sht.Rows
			Sht.Range(j, i, j+1, i).MergeCells = True
		Next

		Sht.Range(j, TotalsRow+1, j+1, Sht.Rows).SetBorder acBrdOutline
	Next

	Sht.Range(1, TotalsRow+1, ColNumber, Sht.Rows).Stripe
	Sht.Range(1, StartNumberOfRows-2, ColNumber, Sht.Rows).Autofit

	Sht.Range(3, TotalsRow, 8, Sht.Rows).Alignment = acRight
	Sht.Range(9, TotalsRow, ColNumber, Sht.Rows).Alignment = acCenter
	Sht.Range(5, TotalsRow+1, ColNumber, Sht.Rows).Font.Bold = True
End Sub

'Тћяюыэџхь яр№рьхђ№шчш№ютрээћщ (яю эхюсѕюфшьюёђш) чря№юё, я№хфёђртыхээћщ ёђ№юъющ SqlQr ---------
Function ExecuteQuery(SqlQr, PrmValues, PrmTypes)
	Dim i
  	Dim Cn, Cmd

	Set Cn = WorkArea.AdoConnection
	Set Cmd = CreateObject("ADODB.Command")
	Set Cmd.ActiveConnection = Cn
    
  Cmd.CommandText = SqlQr
	