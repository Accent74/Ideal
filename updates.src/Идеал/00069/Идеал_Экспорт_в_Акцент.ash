аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ).Value = Rs.Fields(1).Value
		Sheet1.Cell(RowNo, 2).Value = Code.CrCode
		Sheet1.Cell(RowNo, 3).Value = Rs.Fields(0).Value
	
		Rs.MoveNext
	Wend
End Sub
'---
'
'---
Sub ExportToAccent
	Dim xml, i, xmlPackage, xmlTransList, xmlTrans
	Dim RD, xmlFileName, xmlPath

	Set RD = CreateLibObject("Redirect")
	xmlPath = RD.GetPathArray("*.xml")

	If UBound(xmlPath) <> -1 Then
		xmlFileName = xmlPath(0) & "\payroll_data.xml"
	Else
		MsgBox "гърцшђх яѓђќ фыџ xml єрщыют т эрёђ№ющърѕ я№юу№рььћ.", vbExclamation, "Тэшьрэшх !"
		MsgBox "нъёяю№ђ юъюэїшыёџ эхѓфрїэю.", vbInformation, "нъёяю№ђ я№ютюфюъ"
		Exit Sub
	End If

	Set XML = CreateObject("msxml.domdocument")
	Set xmlPackage = XML.CreateElement("package")
	Set xmlTransList = XML.CreateElement("transactions")

	xmlPackage.SetAttribute "timestamp", Now
	xmlPackage.SetAttribute "source-application", "payroll7" 
	xmlPackage.SetAttribute "app-version", App.Version

	For i = START_ROW + 1 To Sheet1.Rows
		If checknull(Sheet1.Cell(i, 3).Value, 0) <> 0 Then
			AppendTrans XML, i, xmlTrans
			xmlTransList.AppendChild xmlTrans
		End If
	Next

	xmlPackage.AppendChild xmlTransList
	XML.AppendChild xmlPackage

	XML.Save xmlFileName
	MsgBox "нъёяю№ђ т єрщы " & xmlFileName & " юъюэїхэ", vbInformation, "нъёяю№ђ я№ютюфюъ"

End Sub
'---
'
'---
Sub AppendTrans(XML, RowNo, ByRef xmlTrans)
	Dim xmlDB, xmlCR, xmlValue

	Set xmlTrans = XML.CreateElement("transaction")
	xmlTrans.SetAttribute "no", RowNo - START_ROW

	Set xmlDB = XML.CreateElement("db")
	xmlDB.SetAttribute "code", checknull(Sheet1.Cell(RowNo, 1).Value, "")
	Set xmlCR = XML.CreateElement("cr")
	xmlCr.SetAttribute "code", checknull(Sheet1.Cell(RowNo, 2).Value, "")
	Set xmlValue = XML.CreateElement("value")
	xmlValue.SetAttribute "sum", checknull(Sheet1.Cell(RowNo, 3).Value, 0)

	xmlTrans.AppendChild xmlDb
	xmlTrans.AppendChild xmlCr
	xmlTrans.AppendChild xmlValue

End Sub
'---
'
'---
Sub ShtBook_OnBarClick(Command)
	ExecBarTag Command
	recalc
End Sub
'---
'
'---

<     ш                Ь Arial ќЈќРќиќ№ќќ ќ      џџџџ                 џџ 
 CSheetPageSheet1Ышёђ 1         61    7                                     f   T   l                                                        џџ  CRow џџ  CCell6      џџ V   @>2>4:8,   M:A?>@B8@C5<K5  2  ?@>3@0<<C  :F5=B                      џџџџ                         џџџџ                            џџ      ?5@8>4                      џџџџ    /=20@L  2 0 1 2   3.                       џџџџ                                   B                               B                    "       
   !C<<0                       џџ       2 3                       џџ       6 5 7                     " џџ    T Ы                           џџ       9 2                       џџ       6 5 7                     " џџ    Y ?                           џџ       9 2                       џџ       6 5 4                     " џџ    вi                            џџ       2 3                       џџ       6 5 7 4                     " џџ    `T                           џџ       6 5 2                       џџ       6 5 3                     " џџ    Cк)                            џџ       9 2                       џџ       6 5 3                     " џџ    A+                            џџ       9 2                       џџ       6 6 1                     " џџ    ќЙЏ                           џџ       6 5 2                       џџ       6 6 1                     " џџ    р~                            џџ       2 3                       џџ       6 6 1                     " џџ     E,                           џџ       6 6 1                       џџ       3 7 7                     " џџ    p                            џџ       6 6 1                       џџ       6 5 7 1                     " џџ    DVK                           џџ       6 6 1                       џџ       6 5 7 3                     " џџ     h                            џџ       6 6 1                       џџ       6 5 7 2                     " џџ    @Ё                            џџ       6 6 1                       џџ       6 4 1 4                     " џџ    №H                      d   d   d   d    <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                 2   2        <     ш               Ь Arial ђш№ѓхьћх т я№юу№рььѓ Ръіхэђ     ццњ                              ццњ    РРР    РРР    РРР    РРР  ццњ                                  рџџ                              џџџ                               џџџџ                             џџ< 
 CBarButton'      нъёяю№ђ т Ръіхэђ0бючфрэшх єрщыр ё я№ютюфърьш фыџ я№юу№рььћ Ръіхэђ    ExportToAccent                                                                       џџџџџџџџџџџџ                                                                        	      ўџџџ§џџџўџџџўџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ   !   џџџџџџџџ    џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Шфхры_нъёяю№ђ_т_Ръіхэђ.ash                                                          66  @   рЯюгZ   @   @`x+YЭ@   СgЗOЭ      Ръіхэђ 7.0                                           џ'#include "definition.avb"

Option Explicit

Const START_ROW = 3

Dim cPeriod
Dim MaxSum


Sub ShtBook_OnLoad
	Setџ'#include "definition.avb"

Option Explicit

Const START_ROW = 3

Dim cPeriod
Dim MaxSum

'---
'
'---
Sub ShtBook_OnLoad
	Set cPeriod = Workarea.Period
	MaxSum = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )

	MakeTransaction
	readonly False
	
End Sub
'---
'
'---
Sub MakeTransaction
	PeriodText = cPeriod.RepTitle
	Sheet1.Rows = START_ROW
	Sheet1.Columns = 3

	CalcSumByDeps "2000", cPeriod.ID, False, False, MaxSum
	CalcSumByDeps "2001", cPeriod.ID, True, False, MaxSum
	CalcSumByDeps "2003", cPeriod.ID, False, True, MaxSum
	CalcSumByDeps "2002", cPeriod.ID, False, False, MaxSum

	CalcByDbCr "92", "661", cPeriod.ID
	CalcByDbCr "652", "661", cPeriod.ID
	CalcByDbCr "23", "661", cPeriod.ID

	CalcSumByCode "2220", cPeriod.ID, "661", "377"
	CalcSumByCode "2410", cPeriod.ID, "661", "6571"
	CalcSumByCode "2420", cPeriod.ID, "661", "6573"
	CalcSumByCode "2430", cPeriod.ID, "661", "6572"
	CalcSumByCode "2400", cPeriod.ID, "661", "6414"

	Sheet1.range(3, START_ROW + 1, 3, Sheet1.Rows).Alignment = acRight
	
	With Sheet1.range(1, START_ROW + 1, Sheet1.Columns, Sheet1.Rows)
		.SetBorder acBrdGrid	
		.Stripe stdBackColor(3), stdBackColor(4)
	End With

	Sheet1.Cell(2, 2).Value = cPeriod.RepTitle
	Sheet1.range(1, 1, Sheet1.Columns, Sheet1.Rows).AutoFit 1 + 2


End Sub
'---
'
'---
Sub CalcSumByCode(CodeCode, pID, DB, CR)
	Dim QD, Rs, RowNo
	Dim Code

	Set Code = Workarea.Code(workarea.GetCodeID(CodeCode))

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS CodeID Long, pID Long ; " & _
					"SELECT Sum(P_JOURNAL.JP_SUM) FROM P_JOURNAL WHERE (((P_JOURNAL.MTC_ID)=[CodeID]) And ((P_JOURNAL.W_PERIOD)=[pID]));"

	QD.Parameters(0).Value = Code.ID
	QD.Parameters(1).Value = pID

	Set Rs = QD.OpenRecordset

	If Not Rs.EOF Then
		RowNo = Sheet1.InsertRow2
		Sheet1.Cell(RowNo, 1).Value = DB
		Sheet1.Cell(RowNo, 2).Value = CR
		Sheet1.Cell(RowNo, 3).Value = Rs.Fields(0).Value
	End If

End Sub
'---
'
'---
Sub CalcByDbCr(Db, Cr, pID)
	Dim QD, Rs, RowNo

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS Db text(255), Cr text(255), pID Long; " & _
					"Select Sum(P_JOURNAL.JP_SUM) AS [Sum-JP_SUM] " & _
					"FROM P_JOURNAL " & _
					"WHERE (((P_JOURNAL.JP_DB)=[Db]) And ((P_JOURNAL.JP_CR)=[Cr]) And ((P_JOURNAL.W_PERIOD)=[pID]));"

	QD.Parameters(0).Value = Db
	QD.Parameters(1).Value = Cr
	QD.Parameters(2).Value = pID

	Set Rs = QD.OpenRecordset

	If Not Rs.EOF Then
		RowNo = Sheet1.InsertR Ь Arial ђш№ѓхьћх т я№юу№рььѓ Ръіхэђ     ццњ                              ццњ    РРР    РРР    РРР    РРР  ццњ                                  рџџ                              џџџ                               џџџџ                             џџ< 
 CBarButton'      нъёяю№ђ т Ръіхэђ0бючфрэшх єрщыр ё я№ютюфърьш фыџ я№юу№рььћ Ръіхэђ    ExportToAccent                                                                       џџџџџџџџџџџџ                                                R o o t   E n t r y                                               џџџџџџџџ                               а#э.YЭ         C o n t e n t s                                                  џџџџ   џџџџ                                       %       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
   ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ               џџџџџџџџ    "   џџџџ    џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Шфхры_нъёяю№ђ_т_Ръіхэђ.ash                                                          67  @   `їЁH^   @   Р8э.YЭ@   СgЗOЭ      Ръіхэђ 7.0                                           џ'#include "definition.avb"

Option Explicit

Const START_ROW = 3

Dim cPeriod
Dim MaxSum


Sub ShtBook_OnLoad
	Setџ'#include "definition.avb"

Option Explicit

Const START_ROW = 3

Dim cPeriod
Dim MaxSum

'---
'
'---
Sub ShtBook_OnLoad
	Set cPeriod = Workarea.Period
	MaxSum = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )

	MakeTransaction
	readonly False
	
End Sub
'---
'
'---
Sub MakeTransaction
	PeriodText = cPeriod.RepTitle
	Sheet1.Rows = START_ROW
	Sheet1.Columns = 3

	CalcSumByDeps "2000", cPeriod.ID, False, False, MaxSum
	CalcSumByDeps "2001", cPeriod.ID, True, False, MaxSum
	CalcSumByDeps "2003", cPeriod.ID, False, True, MaxSum
	CalcSumByDeps "2002", cPeriod.ID, False, False, MaxSum

	CalcByDbCr "92", "661", cPeriod.ID
	CalcByDbCr "652", "661", cPeriod.ID
	CalcByDbCr "23", "661", cPeriod.ID

	CalcSumByCode "2220", cPeriod.ID, "661", "377"
	CalcSumByCode "2410", cPeriod.ID, "661", "6571"
	CalcSumByCode "2420", cPeriod.ID, "661", "6573"
	CalcSumByCode "2430", cPeriod.ID, "661", "6572"
	CalcSumByCode "2400", cPeriod.ID, "661", "6414"

	Sheet1.range(3, START_ROW + 1, 3, Sheet1.Rows).Alignment = acRight
	
	With Sheet1.range(1, START_ROW + 1, Sheet1.Columns, Sheet1.Rows)
		.SetBorder acBrdGrid	
		.Stripe stdBackColor(3), stdBackColor(4)
	End With

	Sheet1.Cell(2, 2).Value = cPeriod.RepTitle
	Sheet1.range(1, 1, Sheet1.Columns, Sheet1.Rows).AutoFit 1 + 2


End Sub
'---
'
'---
Sub CalcSumByCode(CodeCode, pID, DB, CR)
	Dim QD, Rs, RowNo
	Dim Code

	Set Code = Workarea.Code(workarea.GetCodeID(CodeCode))

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS CodeID Long, pID Long ; " & _
					"SELECT Sum(P_JOURNAL.JP_SUM) FROM P_JOURNAL WHERE (((P_JOURNAL.MTC_ID)=[CodeID]) And ((P_JOURNAL.W_PERIOD)=[pID]));"

	QD.Parameters(0).Value = Code.ID
	QD.Parameters(1).Value = pID

	Set Rs = QD.OpenRecordset

	If Not Rs.EOF Then
		RowNo = Sheet1.InsertRow2
		Sheet1.Cell(RowNo, 1).Value = DB
		Sheet1.Cell(RowNo, 2).Value = CR
		Sheet1.Cell(RowNo, 3).Value = Rs.Fields(0).Value
	End If

End Sub
'---
'
'---
Sub CalcByDbCr(Db, Cr, pID)
	Dim QD, Rs, RowNo

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS Db text(255), Cr text(255), pID Long; " & _
					"Select Sum(P_JOURNAL.JP_SUM) AS [Sum-JP_SUM] " & _
					"FROM P_JOURNAL " & _
					"WHERE (((P_JOURNAL.JP_DB)=[Db]) And ((P_JOURNAL.JP_CR)=[Cr]) And ((P_JOURNAL.W_PERIOD)=[pID]));"

	QD.Parameters(0).Value = Db
	QD.Parameters(1).Value = Cr
	QD.Parameters(2).Value = pID

	Set Rs = QD.OpenRecordset

	If Not Rs.EOF Then
		RowNo = Sheet1.InsertRow2
		Sheet1.Cell(RowNo, 1).Value = db
		Sheet1.Cell(RowNo, 2).Value = cr
		Sheet1.Cell(RowNo, 3).Value = Rs.Fields(0).Value
	End If


End Sub
'---
'
'---
Sub CalcSumByDeps(CodeCode, pID, IsDisab, IsAgree, MaxSum)
	Dim QD, Rs, RowNo
	Dim Code

	Set Code = Workarea.Code(workarea.GetCodeID(CodeCode))

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS BaseCode Long, pID Long, IsDisab Long, IsAgree Long, MaxSum Currency, pCent Currency; " & _
					"Select Sum(IIf([JP_SUM]>[MaxSum],[MaxSum]*[pCent],[JP_SUM]*[pCent])), P_JOURNAL.JP_DB, P_JOURNAL.JP_CR " & _
					"FROM AGENTS Right Join (MTD_DEPENDS Left Join (P_JOURNAL Left Join MTD_CODES On P_JOURNAL.MTD_ID = MTD_CODES.MTD_ID) On MTD_DEPENDS.MTD_ID_LEFT = P_JOURNAL.MTC_ID) On AGENTS.AG_ID = P_JOURNAL.AG_ID " & _
					"WHERE (((IIf(IsNull([AG_DISAB]),0,[AG_DISAB]))=[IsDisab]) And ((IIf(IsNull([AG_AGRMNT]),0,[AG_AGRMNT]))=[IsAgree]) And ((P_JOURNAL.W_PERIOD)=[pID]) And ((P_JOURNAL.JP_DONE)=1) And ((MTD_DEPENDS.MTD_ID_RIGHTow2
		Sheet1.Cell(RowNo, 1).Value = db
		Sheet1.Cell(RowNo, 2).Value = cr
		Sheet1.Cell(RowNo, 3).Value = Rs.Fields(0).Value
	End If


End Sub
'---
'
'---
Sub CalcSumByDeps(CodeCode, pID, IsDisab, IsAgree, MaxSum)
	Dim QD, Rs, RowNo
	Dim Code

	Set Code = Workarea.Code(workarea.GetCodeID(CodeCode))

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS BaseCode Long, pID Long, IsDisab Long, IsAgree Long, MaxSum Currency, pCent Currency; " & _
					"Select Sum(IIf([JP_SUM]>[MaxSum],[MaxSum]*[pCent],[JP_SUM]*[pCent])), P_JOURNAL.JP_DB, P_JOURNAL.JP_CR " & _
					"FROM AGENTS Right Join (MTD_DEPENDS Left Join (P_JOURNAL Left Join MTD_CODES On P_JOURNAL.MTD_ID = MTD_CODES.MTD_ID) On MTD_DEPENDS.MTD_ID_LEFT = P_JOURNAL.MTC_ID) On AGENTS.AG_ID = P_JOURNAL.AG_ID " & _
					"WHERE (((IIf(IsNull([AG_DISAB]),0,[AG_DISAB]))=[IsDisab]) And ((IIf(IsNull([AG_AGRMNT]),0,[AG_AGRMNT]))=[IsAgree]) And ((P_JOURNAL.W_PERIOD)=[pID]) And ((P_JOURNAL.JP_DONE)=1) And ((MTD_DEPENDS.MTD_ID_RIGHT)=[BaseCode]) And ((P_JOURNAL.JP_SUM)<>0)) " & _
					"GROUP BY P_JOURNAL.JP_DB, P_JOURNAL.JP_CR " & _
					"HAVING Sum(IIf([JP_SUM]>[MaxSum],[MaxSum]*[pCent],[JP_SUM]*[pCent])) <> 0;"

	QD.Parameters(0).Value = Code.ID
	QD.Parameters(1).Value = pID
	QD.Parameters(2).Value = IsDisab
	QD.Parameters(3).Value = IsAgree
	QD.Parameters(4).Value = MaxSum
	QD.Parameters(5).Value = Code.Percent / 100.0

	Set Rs = QD.OpenRecordset

	While Not Rs.eof
		RowNo = Sheet1.InsertRow2

		Sheet1.Cell(RowNo, 1)=[BaseCode]) And ((P_JOURNAL.JP_SUM)<>0)) " & _
					"GROUP BY P_JOURNAL.JP_DB, P_JOURNAL.JP_CR " & _
					"HAVING Sum(IIf([JP_SUM]>[MaxSum],[MaxSum]*[pCent],[JP_SUM]*[pCent])) <> 0;"

	QD.Parameters(0).Value = Code.ID
	QD.Parameters(1).Value = pID
	QD.Parameters(2).Value = IsDisab
	QD.Parameters(3).Value = IsAgree
	QD.Parameters(4).Value = MaxSum
	QD.Parameters(5).Value = Code.Percent / 100.0

	Set Rs = QD.OpenRecordset

	While Not Rs.eof
		RowNo = Sheet1.InsertRow2

		Sheet1.Cell(RowNo, 1