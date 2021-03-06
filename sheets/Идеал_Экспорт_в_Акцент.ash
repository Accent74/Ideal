аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ+ TrSum
End Sub
'---
'
'---
Sub CalcByDbCr(Db, Cr, pID)
	Dim QD, Rs

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
		AddData DB, CR, Rs.Fields(0).Value
	End If

End Sub
'---
'
'---
Sub CalcSumByDeps(CodeCode, pID, IsDisab, IsAgree, MaxSum)
	Dim QD, Rs
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
		AddData Rs.Fields(1).Value, Code.CrCode, Rs.Fields(0).Value
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
		MsgBox "нъёяю№ђ юъюэїшыёџ эхѓфрїэю.", vbInformation, MSG_BOX_CAPTION
		Exit Sub
	End If

	Set XML = CreateObject("msxml.domdocument")
	Set xmlPackage = XML.CreateElement("package")
	Set xmlTransList = XML.CreateElement("transactions")

	xmlPackage.SetAttribute "timestamp", Now
	xmlPackage.SetAttribute "source-application", "payroll7" 
	xmlPackage.SetAttribute "app-version", App.Version

	For i = START_ROW + 1 To Sheet1.Rows
		AppendTrans XML, i, xmlTrans
		xmlTransList.AppendChild xmlTrans
	Next

	xmlPackage.AppendChild xmlTransList
	XML.AppendChild xmlPackage

	XML.Save xmlFileName '"d:\payroll_data.xml"
	MsgBox "нъёяю№ђ т єрщы " & xmlFileName & " юъюэїхэ", vbInformation, MSG_BOX_CAPTION

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
 CSheetPageSheet1Ышёђ 1         61    7                                     f   T   f                             џџ  CRow џџ  CCell6      џџ V   @>2>4:8,   M:A?>@B8@C5<K5  2  ?@>3@0<<C  :F5=B                      џџџџ                         џџџџ                            џџ      ?5@8>4                      џџџџ    23CAB  2 0 1 4   3.                       џџџџ                                   B                               B                    "       
   !C<<0                       џџ       9 2                       џџ       6 5 1                     " џџ    @ЄЖ                            џџ       9 2                       џџ       6 6 1                     " џџ     lм                           џџ       6 6 1                       џџ       6 5 7 1                     " џџ     /                            џџ       6 6 1                       џџ       6 4 1 4                     " џџ    r                            џџ       6 6 1                       џџ       6 4 2                     " џџ    @~                       d   d   d   d    <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                 2   2        <     ш               Ь Arial ђш№ѓхьћх т я№юу№рььѓ Ръіхэђ     ццњ                              ццњ    РРР    РРР    РРР    РРР  ццњ                                  рџџ                              џџџ                               џџџџ                             џџ< 
 CBarButton'      нъёяю№ђ т Ръіхэђ0бючфрэшх єрщыр ё я№ютюфърьш фыџ я№юу№рььћ Ръіхэђ    ExportToAccent                      Ь Arial ђш№ѓхьћх т я№юу№рььѓ Ръіхэђ     ццњ                              ццњ    РРР    РРР    РРР    РРР  ццњ                                  рџџ                              џџџ                               џџџџ                             џџ< 
 CBarButton'      нъёяю№ђ т Ръіхэђ0бючфрэшх єрщыр ё я№ютюфърьш фыџ я№юу№рььћ Ръіхэђ    ExportToAccent               џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџR o o t   E n t r y                                               џџџџџџџџ                               0!МлвЯ         C o n t e n t s                                                  џџџџ   џџџџ                                       Q)       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
         ўџџџўџџџ§џџџўџџџўџџџ                         џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Шфхры_нъёяю№ђ_т_Ръіхэђ.ash                                                          85  @   їпШ   @   ЭМлвЯ@   СgЗOЭ      Ръіхэђ 7.0                                           џ'#include "definition.avb"

Option Explicit

Const START_ROW = 3

Dim cPeriod
Dim MaxSum


Sub ShtBook_OnLoad
	Setџ'#include "definition.avb"

Option Explicit

Const START_ROW = 3
Const XML_SETTING_FILE_MAME = "payroll_export_settings.xml"
Const MSG_BOX_CAPTION = "нъёяю№ђ фрээћѕ т Ръіхэђ"

Dim cPeriod
Dim MaxSum
Dim dTrans

'---
'
'---
Sub ShtBook_OnLoad
	Dim Rd, xmlFileNamePath, xml

	Set RD = CreateLibObject("Redirect")
	xmlFileNamePath = RD.GetFullPath(XML_SETTING_FILE_MAME)

	If Not xmlFileNamePath = "" Then
		Set cPeriod = Workarea.Period
		MaxSum = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )

		Set dTrans = CreateObject("Scripting.Dictionary")

		Set Xml = LoadXMLSettings(xmlFileNamePath)

		If Not xml Is Nothing Then 
			MakeTransaction xml
		End If
	Else
		MsgBox "Эх эрщфхэ єрщы ё ѓёђрэютърьш §ъёяю№ђр " & XML_SETTING_FILE_MAME, vbExclamation, MSG_BOX_CAPTION 
	End If

	readonly False
	
End Sub
'---
'
'---
Function LoadXMLSettings(xmlFileNamePath)
	Dim xml, strErrText

	Set xml = CreateObject("msxml2.domdocument")

	If xml.Load(xmlFileNamePath) Then
		Set LoadXMLSettings = xml
	Else
		With xml.ParseError
	      strErrText = "Юјшсър чру№ѓчъш XML-фюъѓьхэђр " & vbCrLf & _
  	      "Юјшсър #: " & .errorCode & ": " & .reason & _
  	      "бђ№юър #: " & .Line & vbCrLf & _
  	      "Яючшішџ т ёђ№юъх: " & .linepos & vbCrLf & _
        "Яючшішџ т єрщых: " & .filepos & vbCrLf & _
        "Шёђюїэшъ: " & .srcText & vbCrLf & _
        "Фюъѓьхэђ: " & .url
		End With

		MsgBox strErrText, vbCritical, MSG_BOX_CAPTION 

		Set LoadXMLSettings = Nothing
	End If

End Function
'---
'
'---
Sub MakeTransaction(xml)
	Dim i, j, nodename

	PeriodText = cPeriod.RepTitle
	Sheet1.Rows = START_ROW
	Sheet1.Columns = 3

	With xml.childnodes.Item(0).ChildNodes
		For i = 0 To .Length - 1
			With .Item(i)
				nodename = .nodename

				With .childnodes
					For j = 0 To .Length - 1
						With .Item(j)
							Select Case nodename
								Case "sumbydeps"
									CalcSumByDeps .GetAttribute("code"), cPeriod.ID, CBool2(.GetAttribute("isstolen")),  CBool2(.GetAttribute("isagree")), MaxSum
								Case "sumbydbcr"
									CalcByDbCr .GetAttribute("db"), .GetAttribute("cr"), cPeriod.ID

								Case "sumbycode"
									CalcSumByCode .GetAttribute("code"), cPeriod.ID, .GetAttribute("db"), .GetAttribute("cr")

							End Select
						End With
					Next
				End With
			End With
		Next
	End With

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
Function CBool2(BoolVal)
	If BoolVal = "" Then
		CBool2 = False
	Else
		CBool2 = (LCase(BoolVal) = "true")
	End If
End Function
'---
'
'---
Sub CalcSumByCode(CodeCode, pID, DB, CR)
	Dim QD, Rs
	Dim Code

	Set Code = Workarea.Code(workarea.GetCodeID(CodeCode))

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS CodeID Long, pID Long ; " & _
					"SELECT Sum(P_JOURNAL.JP_SUM) FROM P_JOURNAL WHERE (((P_JOURNAL.MTC_ID)=[CodeID]) And ((P_JOURNAL.W_PERIOD)=[pID]));"

	QD.Parameters(0).Value = Code.ID
	QD.Parameters(1).Value = pID

	Set Rs = QD.OpenRecordset

	If Not Rs.EOF Then
		AddData DB, CR, Rs.Fields(0).Value
	End If

End Sub
'---
'
'---
Sub AddData(DB, CR, TrSum)
	Dim key, RowNo
	
	If checknull(TrSum, 0) = 0 Then Exit Sub

	key = DB & ":" & CR

	If Not dTrans.Exists(DB & ":" & CR) Then
		RowNo = Sheet1.InsertRow2
		Sheet1.Cell(RowNo, 1).Value = DB
		Sheet1.Cell(RowNo, 2).Value = CR

		dTrans.Add key, RowNo
	Else
		RowNo = dTrans(Key)
	End If

	Sheet1.Cell(RowNo, 3).Value = Sheet1.Cell(RowNo, 3).Value + TrSum
End Sub
'---
'
'---
Sub CalcByDbCr(Db, Cr, pID)
	Dim QD, Rs

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS Db text(255), Cr text(255), pID Long; " & _
					"Select Sum(P_JOURNAL.JP_SUM) AS [Sum-JP_R o o t   E n t r y                                               џџџџџџџџ                               а{:qмвЯ         C o n t e n t s                                                  џџџџ   џџџџ                                       П'       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
      ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ          !   "   #   $       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Шфхры_нъёяю№ђ_т_Ръіхэђ.ash                                                          86  @   `щЩ   @   `І/qмвЯ@   СgЗOЭ      Ръіхэђ 7.0                                           џ'#include "definition.avb"

Option Explicit

Const START_ROW = 3

Dim cPeriod
Dim MaxSum


Sub ShtBook_OnLoad
	Setџ'#include "definition.avb"

Option Explicit

Const START_ROW = 3
Const XML_SETTING_FILE_MAME = "payroll_export_settings.xml"
Const MSG_BOX_CAPTION = "нъёяю№ђ фрээћѕ т Ръіхэђ"

Dim cPeriod
Dim MaxSum
Dim dTrans
Dim Pack

'---
'
'---
Sub ShtBook_OnLoad
	Dim Rd, xmlFileNamePath, xml

	Set RD = CreateLibObject("Redirect")
	xmlFileNamePath = RD.GetFullPath(XML_SETTING_FILE_MAME)

	If Not xmlFileNamePath = "" Then
		Set cPeriod = Workarea.Period
		MaxSum = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )

		Set dTrans = CreateObject("Scripting.Dictionary")

		Set Xml = LoadXMLSettings(xmlFileNamePath)

		If Not xml Is Nothing Then 
			MakeTransaction xml
		End If
	Else
		MsgBox "Эх эрщфхэ єрщы ё ѓёђрэютърьш §ъёяю№ђр " & XML_SETTING_FILE_MAME, vbExclamation, MSG_BOX_CAPTION 
	End If

	readonly False
	
End Sub
'---
'
'---
Function LoadXMLSettings(xmlFileNamePath)
	Dim xml, strErrText

	Set xml = CreateObject("msxml2.domdocument")

	If xml.Load(xmlFileNamePath) Then
		Set LoadXMLSettings = xml
	Else
		With xml.ParseError
	      strErrText = "Юјшсър чру№ѓчъш XML-фюъѓьхэђр " & vbCrLf & _
  	      "Юјшсър #: " & .errorCode & ": " & .reason & _
  	      "бђ№юър #: " & .Line & vbCrLf & _
  	      "Яючшішџ т ёђ№юъх: " & .linepos & vbCrLf & _
        "Яючшішџ т єрщых: " & .filepos & vbCrLf & _
        "Шёђюїэшъ: " & .srcText & vbCrLf & _
        "Фюъѓьхэђ: " & .url
		End With

		MsgBox strErrText, vbCritical, MSG_BOX_CAPTION 

		Set LoadXMLSettings = Nothing
	End If

End Function
'---
'
'---
Sub MakeTransaction(xml)
	Dim i, j, nodename

	PeriodText = cPeriod.RepTitle
	Sheet1.Rows = START_ROW
	Sheet1.Columns = 3

	With xml.childnodes.Item(0)
		Pack = CBool2(.GetAttribute("pack"))

		With .ChildNodes
			For i = 0 To .Length - 1
				With .Item(i)
					nodename = .nodename
	
					With .childnodes
						For j = 0 To .Length - 1
							With .Item(j)
								Select Case nodename
									Case "sumbydeps"
										CalcSumByDeps .GetAttribute("code"), cPeriod.ID, CBool2(.GetAttribute("isstolen")),  CBool2(.GetAttribute("isagree")), MaxSum
									Case "sumbydbcr"
										CalcByDbCr .GetAttribute("db"), .GetAttribute("cr"), cPeriod.ID
	
									Case "sumbycode"
										CalcSumByCode .GetAttribute("code"), cPeriod.ID, .GetAttribute("db"), .GetAttribute("cr")
	
								End Select
							End With
						Next
					End With
				End With
			Next
		End With
	End With

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
Function CBool2(BoolVal)
	If BoolVal = "" Then
		CBool2 = False
	Else
		CBool2 = (LCase(BoolVal) = "true")
	End If
End Function
'---
'
'---
Sub CalcSumByCode(CodeCode, pID, DB, CR)
	Dim QD, Rs
	Dim Code

	Set Code = Workarea.Code(workarea.GetCodeID(CodeCode))

	Set QD = Workarea.DAODAtabase.CreateQueryDef("")
	QD.SQL = 	"PARAMETERS CodeID Long, pID Long ; " & _
					"SELECT Sum(P_JOURNAL.JP_SUM) FROM P_JOURNAL WHERE (((P_JOURNAL.MTC_ID)=[CodeID]) And ((P_JOURNAL.W_PERIOD)=[pID]));"

	QD.Parameters(0).Value = Code.ID
	QD.Parameters(1).Value = pID

	Set Rs = QD.OpenRecordset

	If Not Rs.EOF Then
		AddData DB, CR, Rs.Fields(0).Value
	End If

End Sub
'---
'
'---
Sub AddData(DB, CR, TrSum)
	Dim key, RowNo
	
	If Pack Then
		If checknull(TrSum, 0) = 0 Then Exit Sub
	
		key = DB & ":" & CR
	
		If Not dTrans.Exists(DB & ":" & CR) Then
			RowNo = Sheet1.InsertRow2
			Sheet1.Cell(RowNo, 1).Value = DB
			Sheet1.Cell(RowNo, 2).Value = CR
	
			dTrans.Add key, RowNo
		Else
			RowNo = dTrans(Key)
		End If
	Else
		RowNo = Sheet1.InsertRow2
		Sheet1.Cell(RowNo, 1).Value = DB
		Sheet1.Cell(RowNo, 2).Value = CR		
	End If

	Sheet1.Cell(RowNo, 3).Value = Sheet1.Cell(RowNo, 3).Value 