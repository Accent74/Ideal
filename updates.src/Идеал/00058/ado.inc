' ---------------------------------------------
' Константы Microsoft Active Data Objects (ADO)
' Version 2.0
' 29 Jun 2000
' ---------------------------------------------

' Типы данных для полей и параметров
Const adEmpty            =   0 ' vbEmpty
Const adTinyInt          =  16
Const adSmallInt         =   2 ' vbInteger
Const adInteger          =   3 ' vbLong
Const adBigInt           =  20
Const adUnsignedTinyInt  =  17 ' vbByte
Const adUnsignedSmallInt =  18
Const adUnsignedInt      =  19
Const adUnsignedBigInt   =  21
Const adSingle           =   4 ' vbSingle
Const adDouble           =   5 ' vbDouble
Const adCurrency         =   6 ' vbCurrency
Const adDecimal          =  14 ' vbDecimal
Const adNumeric          = 131
Const adBoolean				   =  11 ' vbBoolean
Const adError            =  10 ' vbError
Const adUserDefined      = 132
Const adVariant          =  12 ' vbVariant
Const adIDispatch        =   9 ' vbObject
Const adIUnknown         =  13
Const adGUID             =  72
Const adDate             =   7 ' vbDate
Const adDBDate           = 133
Const adDBTime           = 134
Const adDBTimeStamp      = 135
Const adBSTR             =   8 ' vbString
Const adChar             = 129
Const adVarChar          = 200
Const adLongVarChar      = 201
Const adWChar            = 130
Const adVarWChar         = 202
Const adLongVarWChar     = 203
Const adBinary           = 128
Const adVarBinary        = 204
Const adLongVarBinary    = 205
Const adChapter          = 136
Const adFileTime         =  64
Const adDBFileTime       = 137
Const adPropVariant      = 138
Const adVarNumeric       = 139

' Типы курсоров
Const adOpenUnspecified  = -1
Const adOpenForwardOnly  = 0
Const adOpenKeyset       = 1
Const adOpenDynamic      = 2
Const adOpenStatic       = 3

' Где открывать курсор
Const adUseServer        = 2
Const adUseClient        = 3

' Виды блокировок
Const adLockUnspecified     = -1
Const adLockReadOnly        = 1
Const adLockPessimistic     = 2
Const adLockOptimistic      = 3
Const adLockBatchOptimistic = 4

' Типы команд
Const adCmdUnknown    = 8 ' автоопределение (по умолчанию)
Const adCmdText       = 1 ' инструкция SQL
Const adCmdTable      = 2 ' имя таблицы
Const adCmdStoredProc = 4 ' вызов хранимой процедуры

' "Направление" передачи параметров
Const adParamUnknown     = 0
Const adParamInput       = 1 ' вход
Const adParamOutput      = 2 ' выход
Const adParamInputOutput = 3 ' вход-выход
Const adParamReturnValue = 4 ' возвращаемое значение

