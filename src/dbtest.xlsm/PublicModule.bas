Attribute VB_Name = "PublicModule"
Option Explicit
'Option Base 1
'参照設定
'Microsoft AciteX Data Objects 2.8 Library      %ProgramFiles(x86)%\Common Files\System\msado28.tlb
'Microsoft ADO Ext. 6.0 for DDL and Security    %ProgramFiles(x86)%\Common Files\System\msadox.dll
'Microsoft Scripting Runtime                    %SystemRoot%\SysWOW64\scrrun.dll
'Microsoft DAO 3.6 Object Library               %ProgramFiles(x86)%\Common Files\Microfost Shared\DAO\dao360.dll
'UNC対応のため、Win32API使用
Public Declare PtrSafe Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As LongPtr
'日付をミリ秒単位で取得するのにWin32APIを使用
'SYSTEMTIME構造体定義
Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
'関数定義
Public Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Function isUnicodePath(ByVal strCurrentPath As String) As Boolean
    'パス名にUnicodeが含まれていればTrueを返し、イベント無効にする（マクロ実行しずらいよね）
    Dim strSJIS As String           'パス名を一旦SJISに変換したもの
    Dim strReUnicode As String      'SJISに変換したパス名を再度Unicodeにしたもの
    strSJIS = StrConv(strCurrentPath, vbFromUnicode)
    strReUnicode = StrConv(strSJIS, vbUnicode)
    If strReUnicode <> strCurrentPath Then
        'うにこーどとSJIS変換して戻ってきたのが違う→Unicodeあり
        isUnicodePath = True
        Exit Function
    Else
        '同じなのでうにこーどなし
        isUnicodePath = False
        Exit Function
    End If
End Function
Public Function IsDBFileExist() As Boolean
    Dim fsoObj As FileSystemObject
    Set fsoObj = New FileSystemObject
    'DBファイルの有無を確認する
    ChCurrentToDBDirectory
    If Not fsoObj.FileExists(constJobNumberDBname) Then
        MsgBox "DBファイルが見つからないようなので新規作成します"
        InitialDBCreate
    End If
End Function
Public Function ChCurrentDirW(ByVal DirName As String)
    'UNICODE対応ChCurrentDir
    'SetCurrentDirectoryW（UNICODE）なので
    'StrPtrを介す必要がある・・？
    SetCurrentDirectoryW StrPtr(DirName)
End Function
Public Sub CheckInitialTableJSON()
    '初期テーブル作成用のJSONがあるか確認する
    Dim fsoJSON As FileSystemObject
    Set fsoJSON = New FileSystemObject
    Call ChCurrentToDBDirectory
    If Not fsoJSON.FileExists(JSON_File_InitialDB) Then
        MsgBox "初期テーブル作成用JSONが見つからないため作成します"
        Debug.Print "何故か初期テーブル作成用JSONが見つからない、作成"
        Call CreateInitialTableJSON
    End If
End Sub
Public Sub CreateInitialTableJSON()
    '初期テーブル作成用JSON作成
    Dim dicJSONObject As Dictionary
    Dim strJSON As String
    Dim streamInitialJSON As ADODB.Stream
    Dim sqlbInitial As clsSQLStringBuilder
    Dim strSQLJsonTable As String
    Dim strSQLKishu As String
    Dim strSQLLog As String
    Set dicJSONObject = New Dictionary
    Set streamInitialJSON = New ADODB.Stream
    Set sqlbInitial = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    'カレントをDBディレクトリに移動
    Call ChCurrentToDBDirectory
    'JSONテーブルSQL
    strSQLJsonTable = strSQLJsonTable & strTable1_NextTable & sqlbInitial.addQuote(Table_JSON) & strTable2_Next1stField
    strSQLJsonTable = strSQLJsonTable & sqlbInitial.addQuote(JSON_Field_Name) & strTable3_TEXT & strTable_NotNull & strTable_Unique & strTable4_EndRow
    strSQLJsonTable = strSQLJsonTable & sqlbInitial.addQuote(JSON_Field_string) & strTable3_JSON & strTable_NotNull & strTable4_EndRow
    strSQLJsonTable = strSQLJsonTable & sqlbInitial.addQuote(Field_Initialdate) & strTable3_TEXT & strTable_Default & "CURRENT_TIMESTAMP" & strTable4_EndRow
    strSQLJsonTable = strSQLJsonTable & sqlbInitial.addQuote(Field_Update) & strTable3_TEXT & strTable5_EndSQL
    '機種テーブルSQL
    strSQLKishu = strSQLKishu & strTable1_NextTable & sqlbInitial.addQuote(Table_Kishu) & strTable2_Next1stField
    strSQLKishu = strSQLKishu & sqlbInitial.addQuote(Kishu_Header) & strTable3_TEXT & strTable_NotNull & strTable_Unique & strTable4_EndRow
    strSQLKishu = strSQLKishu & sqlbInitial.addQuote(Kishu_KishuName) & strTable3_TEXT & strTable_NotNull & strTable_Unique & strTable4_EndRow
    strSQLKishu = strSQLKishu & sqlbInitial.addQuote(Kishu_KishuNickname) & strTable3_TEXT & strTable_NotNull & strTable_Unique & strTable4_EndRow
    strSQLKishu = strSQLKishu & sqlbInitial.addQuote(Kishu_TotalKeta) & strTable3_NUMERIC & strTable_NotNull & strTable4_EndRow
    strSQLKishu = strSQLKishu & sqlbInitial.addQuote(Kishu_RenbanKetasuu) & strTable3_NUMERIC & strTable_NotNull & strTable4_EndRow
    strSQLKishu = strSQLKishu & sqlbInitial.addQuote(Field_Initialdate) & strTable3_TEXT & strTable_Default & "CURRENT_TIMESTAMP" & strTable4_EndRow
    strSQLKishu = strSQLKishu & sqlbInitial.addQuote(Field_Update) & strTable3_TEXT & strTable5_EndSQL
    'JSONテーブル
    dicJSONObject.Add Table_JSON, New Dictionary
    dicJSONObject(Table_JSON).Add JSON_Table_SQL, strSQLJsonTable       'JSONテーブル作成用SQL
    dicJSONObject(Table_JSON).Add JSON_Table_Description, "JSON情報格納テーブル"
    '機種テーブル
    dicJSONObject.Add Table_Kishu, New Dictionary
    dicJSONObject(Table_Kishu).Add JSON_Table_SQL, strSQLKishu              'テーブル作成用SQL格納
    dicJSONObject(Table_Kishu).Add JSON_Table_Description, "機種別情報格納テーブル、機種ヘッダ、履歴桁数・枚 per シートの情報等"
    dicJSONObject(Table_Kishu).Add JSON_AppendField, New Dictionary
    dicJSONObject(Table_Kishu)(JSON_AppendField).Add Kishu_Header, New Dictionary
'    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_Header).Add JSON_Table_SQL           '機種ヘッダ作成用SQL格納
    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_Header).Add JSON_Table_Description, "機種判別ヘッダフィールド UNIQUE、NOT NULL制約"
    dicJSONObject(Table_Kishu)(JSON_AppendField).Add Kishu_KishuName, New Dictionary
'    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_KishuName).Add JSON_Table_SQL             '機種名SQL
    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_KishuName).Add JSON_Table_Description, "機種名フィールド、原則制作指示書の図番 UNIQUE"
    dicJSONObject(Table_Kishu)(JSON_AppendField).Add Kishu_KishuNickname, New Dictionary
'    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_KishuNickname).Add JSON_Table_SQL    '機種通称名SQL
    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_KishuNickname).Add JSON_Table_Description, "機種通称名、シート名やコンボボックスの項目に使う、日本語OK UNIQUE"
    dicJSONObject(Table_Kishu)(JSON_AppendField).Add Kishu_TotalKeta, New Dictionary
'    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_TotalKeta).Add JSON_Table_SQL        '機種トータルSQL
    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_TotalKeta).Add JSON_Table_Description, "機種の履歴のトータル桁数 NUMERIC"
    dicJSONObject(Table_Kishu)(JSON_AppendField).Add Kishu_RenbanKetasuu, New Dictionary
'    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_RenbanKetasuu).Add JSON_Table_SQL    '機種連番桁数SQL
    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_RenbanKetasuu).Add JSON_Table_Description, "機種の連番部分の桁数 NUMERIC"
    dicJSONObject(Table_Kishu)(JSON_AppendField).Add Kishu_Mai_Per_Sheet, New Dictionary
'    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_Mai_Per_Sheet).Add JSON_Table_SQL    '機種、mai per sheetSQL
    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_Mai_Per_Sheet).Add JSON_Table_Description, "1シートあたりの枚数 NUMERIC"
    dicJSONObject(Table_Kishu)(JSON_AppendField).Add Kishu_Barcord_Read_Number, New Dictionary
'    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_Barcord_Read_Number).Add JSON_Table_SQL  '機種、バーコード読み取り数SQL
    dicJSONObject(Table_Kishu)(JSON_AppendField)(Kishu_Barcord_Read_Number).Add JSON_Table_Description, "1シートあたりバーコードの数 NUMERIC"
    strJSON = JsonConverter.ConvertToJson(dicJSONObject)
    GoTo CloseAndExit
    Exit Sub
CloseAndExit:
    Set dicJSONObject = Nothing
    Set streamInitialJSON = Nothing
    Set sqlbInitial = Nothing
    Exit Sub
ErrorCatch:
    Debug.Print "CreateInitialJSON code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
    Exit Sub
End Sub
Public Function InitialDBCreate() As Boolean
    Dim isCollect As Boolean
    Dim strSQL  As String
    Dim dbSQLite3 As clsSQLiteHandle
    Set dbSQLite3 = New clsSQLiteHandle
    On Error GoTo ErrorCatch
    'DBディレクトリへ移動
    Call ChCurrentToDBDirectory
    '初期テーブル作成用SQL文作成(T_Kishu)
    strSQL = ""
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_Kishu & """ (" & vbCrLf & """"
    strSQL = strSQL & Kishu_Header & """ TEXT NOT NULL UNIQUE," & vbCrLf & """"
    strSQL = strSQL & Kishu_KishuName & """ TEXT NOT NULL UNIQUE," & vbCrLf & """"
    strSQL = strSQL & Kishu_KishuNickname & """ TEXT NOT NULL UNIQUE," & vbCrLf & """"
    strSQL = strSQL & Kishu_TotalKeta & """ NUMERIC NOT NULL," & vbCrLf & """"
    strSQL = strSQL & Kishu_RenbanKetasuu & """ NUMERIC NOT NULL," & vbCrLf & """"
    strSQL = strSQL & Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP," & vbCrLf & """"
    strSQL = strSQL & Field_Update & """ TEXT)"
    isCollect = dbSQLite3.DoSQL_No_Transaction(strSQL)
    Set dbSQLite3 = Nothing
'    'テスト実装_SQL作成テスト
'    isCollect = CreateTable_by_KishuName("Test15")
'    If Not isCollect Then
'        InitialDBCreate = False
'    End If
    '正常終了
    InitialDBCreate = True
    Exit Function
ErrorCatch:
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
    End If
    Exit Function
End Function
Public Function registNewKishu_to_KishuTable(ByVal strKishuheader As String, ByVal strKishuname As String, ByVal strKishuNickname As String, _
                                ByVal byteTotalKetasu As Byte, ByVal byteRenbanKetasu As Byte) As Boolean
    '新機種登録動作
    Dim dbSQLite3 As clsSQLiteHandle
    Set dbSQLite3 = New clsSQLiteHandle
    Dim strSQLlocal As String
    Dim isCollect As Boolean
    'テーブルの有無確認
    If Not IsTableExist(Table_Kishu) Then
        MsgBox "機種テーブルが無かったので追加します"
        InitialDBCreate
    End If
    'SQL組み立て
    strSQLlocal = "INSERT INTO " & Table_Kishu & _
                    " (" & Kishu_Header & "," & Kishu_KishuName & "," & Kishu_KishuNickname & _
                    "," & Kishu_TotalKeta & "," & Kishu_RenbanKetasuu & _
                    ") VALUES (""" & strKishuheader & """,""" & strKishuname & """,""" & strKishuNickname & """," & _
                    byteTotalKetasu & "," & byteRenbanKetasu & ")"
    dbSQLite3.SQL = strSQLlocal
    isCollect = dbSQLite3.DoSQL_No_Transaction()
    Set dbSQLite3 = Nothing
    If Not isCollect Then
        MsgBox "機種テーブル追加中にエラー発生"
        Debug.Print Err.Description
        registNewKishu_to_KishuTable = False
        Exit Function
    End If
    '続いて機種名より機種名依存のテーブルを作成していく
    isCollect = CreateTable_by_KishuName(strKishuname)
    If Not isCollect Then
        MsgBox "機種別テーブル追加中にエラー"
        registNewKishu_to_KishuTable = False
        Exit Function
    End If
    MsgBox "機種追加完了"
    registNewKishu_to_KishuTable = True
End Function
Public Sub ChCurrentToDBDirectory()
    'カレントディレクトリをDBディレクトリに移動する
    'カレントディレクトリの取得（UNCパス対応）
    Dim strCurrentDir As String
    Dim fso As New scripting.FileSystemObject
'    'カレントディレクトリをブックのディレクトリに変更
'    ChCurrentDirW (ThisWorkbook.Path)
'    strCurrentDir = CurDir
    'DataBaseディレクトリの存在有無確認"
    If fso.FolderExists(constDatabasePath) <> True Then
        'ディレクトリ存在しない場合作成しよ？
        MsgBox "データベースフォルダが無いため作成します。"
        MkDir constDatabasePath
    End If
    'データベースディレクトリに移動
    strCurrentDir = constDatabasePath
    ChCurrentDirW (strCurrentDir)
End Sub
Public Function GetNameRange(ByVal strSerchName As String, Optional ByRef shTarget As Worksheet)
    '名前定義に指定されたものが存在するか調べて、存在したら名前定義そのものを返す
    Dim rngNameLocal As Name
    Dim strSerchParentName As String
    '通常はActiveSheetがParentだが、対象が指定されていた場合はその名前を使用
    If shTarget Is Nothing Then
        '指定されない場合→Activesheet
        strSerchParentName = ActiveSheet.Name
    Else
        '指定されている場合はその名前で
        strSerchParentName = shTarget.Name
    End If
    For Each rngNameLocal In ActiveWorkbook.Names
        If rngNameLocal.RefersToRange.Parent.Name = strSerchParentName And rngNameLocal.Name = strSerchName Then
            'シート名(既定)及び名前が一致
            '一致した名前をそのまま返してあげて・・・
            Set GetNameRange = rngNameLocal
            Set rngNameLocal = Nothing
            Exit Function
        End If
    Next rngNameLocal
    'ここまで出てきたって事は無いのよ・・・
    Set rngNameLocal = Nothing
    GetNameRange = errcxlNameNotFound
End Function
Public Function CreateTable_by_KishuName(ByRef strKishuname As String) As String
    '機種名を引数として取り、それを元に機種名依存のテーブルを作成するSQL文を返す
    Dim strSQL As String
    Dim adstreamReader As ADODB.Stream
    Dim isCollect As Boolean
    Dim dbSQLite3 As clsSQLiteHandle
    Set dbSQLite3 = New clsSQLiteHandle
    'T_Jobdata ジョブの履歴とJob番号
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_JobDataPri & strKishuname & """ (" & vbCrLf
    strSQL = strSQL & """" & Job_Number & """ TEXT NOT NULL," & vbCrLf
    strSQL = strSQL & """" & Job_RirekiHeader & """ TEXT NOT NULL," & vbCrLf
    strSQL = strSQL & """" & Job_RirekiNumber & """ NUMERIC NOT NULL UNIQUE," & vbCrLf
    strSQL = strSQL & """" & Job_Rireki & """ TEXT NOT NULL UNIQUE," & vbCrLf
    strSQL = strSQL & """" & Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP," & vbCrLf
    strSQL = strSQL & """" & Field_Update & """ TEXT," & vbCrLf
    strSQL = strSQL & "Primary Key(""" & Job_Rireki & """)" & vbCrLf
    strSQL = strSQL & ");"
    isCollect = dbSQLite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    strSQL = ""
    '続いてT_Barcorde バーコードテーブル（ピッてするやつ）
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_Barcodepri & strKishuname & """ (" & vbCrLf
    strSQL = strSQL & """" & Field_BarcordeNumber & """ TEXT NOT NULL," & vbCrLf
    strSQL = strSQL & """" & Laser_Rireki & """ TEXT NOT NULL UNIQUE," & vbCrLf
    strSQL = strSQL & """" & Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP," & vbCrLf
    strSQL = strSQL & """" & Field_Update & """ TEXT," & vbCrLf
    strSQL = strSQL & "Primary Key(""" & Laser_Rireki & """)" & vbCrLf
    strSQL = strSQL & ");"
    isCollect = dbSQLite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    strSQL = ""
    '最後にリトライ履歴（いるの？これ）
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_Retrypri & strKishuname & """ (" & vbCrLf
    strSQL = strSQL & """" & Field_BarcordeNumber & """ TEXT NOT NULL," & vbCrLf
    strSQL = strSQL & """" & Laser_Rireki & """ TEXT NOT NULL," & vbCrLf
    strSQL = strSQL & """" & Retry_Reason & """ TEXT," & vbCrLf
    strSQL = strSQL & """" & Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP," & vbCrLf
    strSQL = strSQL & """" & Field_Update & """ TEXT" & vbCrLf
    strSQL = strSQL & ");"
    isCollect = dbSQLite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    'テーブル追加は全部終わった
    'インデックス作成
    'バーコードテーブル
    strSQL = ""
    strSQL = "CREATE UNIQUE INDEX IF NOT EXISTS ""ix" & Table_Barcodepri & strKishuname & """ ON """ & _
    Table_Barcodepri & strKishuname & """ (""" & Laser_Rireki & """ ASC);"
    isCollect = dbSQLite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    'ジョブ履歴テーブル
    strSQL = ""
    strSQL = "CREATE UNIQUE INDEX IF NOT EXISTS ""ix" & Table_JobDataPri & strKishuname & """ ON """ & _
    Table_JobDataPri & strKishuname & """ (""" & Job_Rireki & """ ASC);"
    isCollect = dbSQLite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    'インデックス作成終了
    Set dbSQLite3 = Nothing
    '追加フィールド更新へ
    Call CheckNewField
    CreateTable_by_KishuName = True
    Exit Function
End Function
Public Sub CheckNewField()
    Dim dbTableAdd As clsSQLiteHandle
    Dim strSQL As String
    Dim varTableList As Variant
    Dim intTableCounter As Integer
    Dim arrStr_Kishu_AppendField() As String
    Dim arrStr_Kishu_Type() As String
    Dim arrStr_Job_AppendField() As String
    Dim arrStr_Job_Type() As String
    Dim arrStr_BarCorde_AppendField() As String
    Dim arrStr_BarCorde_Type() As String
    Dim arrStr_Retry_AppendField() As String
    Dim arrStr_Retry_Type() As String
    Dim arrStr_Index_AppendField() As String
    On Error GoTo ErrorCatch
    '追加フィールド定義
    arrStr_Kishu_AppendField = Split(Kishu_Mai_Per_Sheet & "," & Kishu_Barcord_Read_Number, ",")
    arrStr_Kishu_Type = Split("NUMERIC" & "," & "NUMERIC", ",")
    arrStr_Job_AppendField = Split(Job_KanbanChr & "," & Job_ProductDate & "," & Field_LocalInput & "," & Field_RemoteInput & "," & Job_KanbanNumber, ",")
    arrStr_Job_Type = Split("TEXT" & "," & "TEXT" & "," & "NUMERIC" & "," & "NUMERIC" & "," & "NUMERIC", ",")
    arrStr_BarCorde_AppendField = Split(Field_LocalInput & "," & Field_RemoteInput, ",")
    arrStr_BarCorde_Type = Split("NUMERIC" & "," & "NUMERIC", ",")
    arrStr_Retry_AppendField = Split(Field_LocalInput & "," & Field_RemoteInput, ",")
    arrStr_Retry_Type = Split("NUMERIC" & "," & "NUMERIC", ",")
    arrStr_Index_AppendField = Split(Job_Number & "," & Field_Initialdate, ",")
    Set dbTableAdd = New clsSQLiteHandle
    'テーブルとかフィールドとかがんがん追加するやつ
    If Not IsTableExist(Table_Kishu) Then
        Call InitialDBCreate
    End If
    'ログテーブルを追加してやる
    strSQL = ""
    strSQL = strSQL & strOLDAddTable1_NextTable & Table_Log & strOLDAddTable2_Field1_Next_Field & Log_ActionType
    strSQL = strSQL & strOLDAddTable_TEXT_Next_Field & Log_Table
    strSQL = strSQL & strOLDAddTable_TEXT_Next_Field & Log_StartRireki
    strSQL = strSQL & strOLDAddTable_TEXT_Next_Field & Log_Maisuu
    strSQL = strSQL & strOLDAddTable_NUMELIC_Next_Field & Log_JobNumber
    strSQL = strSQL & strOLDAddTable_TEXT_Next_Field & Log_RirekiHeader
    strSQL = strSQL & strOLDAddTable_TEXT_Next_Field & Log_BarcordNumber
    strSQL = strSQL & strOLDAddTable_TEXT_Next_Field & Log_SQL
    strSQL = strSQL & strOLDAddTable_TEXT_Next_Field & Field_LocalInput
    strSQL = strSQL & strOLDAddTable_NUMELIC_Next_Field & Field_RemoteInput & strOLDAddTable_Numeric_Last
    dbTableAdd.SQL = strSQL
    Call dbTableAdd.DoSQL_No_Transaction
    Set dbTableAdd = Nothing
    'テーブル一覧を受け取る
    Set dbTableAdd = New clsSQLiteHandle
    dbTableAdd.SQL = "select name from sqlite_master where type = ""table"";"
    Call dbTableAdd.DoSQL_No_Transaction
    varTableList = dbTableAdd.RS_Array
    Set dbTableAdd = Nothing
    'テーブル数分ループ
    For intTableCounter = LBound(varTableList, 1) To UBound(varTableList, 1)
        If Mid(varTableList(intTableCounter, 0), 1, Len(Table_Kishu)) = Table_Kishu Then
            '機種テーブル
            'フィールド追加
            Call AppendFieldbyTableName(varTableList(intTableCounter, 0), arrStr_Kishu_AppendField, arrStr_Kishu_Type)
        ElseIf Mid(varTableList(intTableCounter, 0), 1, Len(Table_JobDataPri)) = Table_JobDataPri Then
            'Jobテーブル
            'フィールド追加
            Call AppendFieldbyTableName(varTableList(intTableCounter, 0), arrStr_Job_AppendField, arrStr_Job_Type)
            'Index追加
            Call AppendIndexbyTableName(varTableList(intTableCounter, 0), arrStr_Index_AppendField)
        ElseIf Mid(varTableList(intTableCounter, 0), 1, Len(Table_Barcodepri)) = Table_Barcodepri Then
            'バーコードテーブル
            'フィールド追加
            Call AppendFieldbyTableName(varTableList(intTableCounter, 0), arrStr_BarCorde_AppendField, arrStr_BarCorde_Type)
        ElseIf Mid(varTableList(intTableCounter, 0), 1, Len(Table_Retrypri)) = Table_Retrypri Then
            'リトライテーブル
            'フィールド追加
            Call AppendFieldbyTableName(varTableList(intTableCounter, 0), arrStr_Retry_AppendField, arrStr_Retry_Type)
        Else
            Debug.Print "よくわからないテーブルだった"
        End If
    Next intTableCounter
ErrorCatch:
    Set dbTableAdd = Nothing
    Debug.Print "AppendField code: " & Err.Number & "Description " & Err.Description
    Exit Sub
CloseAndExit:
    Set dbTableAdd = Nothing
    Exit Sub
End Sub
Public Sub AppendFieldbyTableName(ByVal strargTableName As String, ByRef arrargstrField() As String, ByRef arrargstrType() As String)
    Dim dbAppendField As clsSQLiteHandle
    Dim byteFieldCounter As Byte
    Dim strSQL As String
    For byteFieldCounter = LBound(arrargstrField) To UBound(arrargstrField)
        If Not IsFieldExist(strargTableName, arrargstrField(byteFieldCounter)) Then
            'フィールドが無いようなので、追加に入る
            Set dbAppendField = New clsSQLiteHandle
            strSQL = ""
            strSQL = strSQL & strOLDAddField1_NextTableName & strargTableName
            strSQL = strSQL & strOLDAddField2_NextFieldName & arrargstrField(byteFieldCounter)
            If arrargstrType(byteFieldCounter) = "NUMERIC" Then
                'NUMERICの場合
                strSQL = strSQL & strOLDAddField3_Numeric_Last
            ElseIf arrargstrType(byteFieldCounter) = "TEXT" Then
                'TEXTの場合
                strSQL = strSQL & strOLDAddField3_Text_Last
            ElseIf arrargstrType(byteFieldCounter) = "JSON" Then
                'JSONの場合
            End If
            dbAppendField.SQL = strSQL
            Call dbAppendField.DoSQL_No_Transaction
            Set dbAppendField = Nothing
        End If
    Next byteFieldCounter
End Sub
Public Function IsFieldExist(ByVal strargTableName As String, ByVal strargFieldName As String) As Boolean
    '特定のテーブルに指定されたフィールドがあるかどうか
    Dim dbIsField As clsSQLiteHandle
    Dim varReturnValue As Variant
    Dim bytFieldCounter As Byte
    'テーブルのフィールド名一覧を取得
    Set dbIsField = New clsSQLiteHandle
    dbIsField.SQL = "select name from pragma_table_info(""" & strargTableName & """);"
    Call dbIsField.DoSQL_No_Transaction
    varReturnValue = dbIsField.RS_Array
    Set dbIsField = Nothing
    'フィールド数分ループ
    For bytFieldCounter = LBound(varReturnValue, 1) To UBound(varReturnValue, 1)
        If varReturnValue(bytFieldCounter, 0) = strargFieldName Then
            IsFieldExist = True
            Exit Function
        End If
    Next bytFieldCounter
    IsFieldExist = False
    Exit Function
End Function
Public Sub AppendIndexbyTableName(ByVal strargTableName As String, arrstrargField() As String)
    Dim dbIndexAdd As clsSQLiteHandle
    Dim byteFieldCounter As Byte
    Dim strSQL As String
    byteFieldCounter = LBound(arrstrargField)
    Do While byteFieldCounter <= UBound(arrstrargField)
        If byteFieldCounter = LBound(arrstrargField) Then
            '初回のみ
            strSQL = ""
            strSQL = strSQL & strOLDIndex1_NextTable & strargTableName
            strSQL = strSQL & strOLDIndex2_NextTable & strargTableName
            strSQL = strSQL & strOLDIndex3_Field1 & arrstrargField(byteFieldCounter)
        End If
        byteFieldCounter = byteFieldCounter + 1
        If byteFieldCounter > UBound(arrstrargField) Then
            'ここは最後に来るところ
            strSQL = strSQL & strOLDIndex5_Last
        Else
            '途中
            strSQL = strSQL & strOLDIndex4_FieldNext & arrstrargField(byteFieldCounter)
        End If
    Loop
    Set dbIndexAdd = New clsSQLiteHandle
    dbIndexAdd.SQL = strSQL
    Call dbIndexAdd.DoSQL_No_Transaction
    Set dbIndexAdd = Nothing
End Sub
Public Function getArryDimmensions(ByRef varArry As Variant) As Byte
    '配列の次元数を返す（Byteまでしか対応しないよ）
    Dim byteLocalCounter As Byte
    Dim longRows As Long
    If Not IsArray(varArry) Then
        MsgBox ("配列じゃないっぽいのが来たので中止です")
        getArryDimmensions = False
        Exit Function
    End If
    byteLocalCounter = 0
    On Error Resume Next
    Do While Err.Number = 0
        byteLocalCounter = byteLocalCounter + 1
        longRows = UBound(varArry, byteLocalCounter)
    Loop
    byteLocalCounter = byteLocalCounter - 1
    Err.Clear
    getArryDimmensions = byteLocalCounter
    Exit Function
 End Function
Public Function getKishuInfoByRireki(strargRireki As String) As typKishuInfo
    '履歴を元に機種情報を返す
    '返り値はKishuInfo型（ユーザー定義構造体）
    'ぐろばんる変数にあるのを使うようになりました
    Dim Kishu As typKishuInfo
    Dim longKishuCounter As Long
    On Error GoTo ErrorCatch
    If strargRireki = "" Then
        MsgBox "機種情報検索には履歴が必須です"
        getKishuInfoByRireki = Kishu
        GoTo CloseAndExit
    End If
    'ぐろばんるのが初期化されているかチェック
    If (Not arrKishuInfoGlobal) = -1 Then
        'ここに来ると未初期化らしい・・・
        Call GetAllKishuInfo_Array
    End If
    If boolNoTableKishuRecord = True Then
        '機種テーブルが空の場合は機種登録画面を表示
        strRegistRireki = strargRireki
        frmRegistNewKishu.Show
    End If
Serch_From_GlobalKishuList:
    For longKishuCounter = LBound(arrKishuInfoGlobal, 1) To UBound(arrKishuInfoGlobal, 1)
        If Mid(strargRireki, 1, Len(arrKishuInfoGlobal(longKishuCounter).KishuHeader)) = _
            arrKishuInfoGlobal(longKishuCounter).KishuHeader Then
            '機種ヘッダが一致したので、KishuInfoを返して終了
            '機種登録OKフラグを立てる
            boolRegistOK = True
            Kishu.KishuHeader = arrKishuInfoGlobal(longKishuCounter).KishuHeader
            Kishu.KishuName = arrKishuInfoGlobal(longKishuCounter).KishuName
            Kishu.KishuNickName = arrKishuInfoGlobal(longKishuCounter).KishuNickName
            Kishu.TotalRirekiketa = arrKishuInfoGlobal(longKishuCounter).TotalRirekiketa
            Kishu.RenbanKetasuu = arrKishuInfoGlobal(longKishuCounter).RenbanKetasuu
            getKishuInfoByRireki = Kishu
            GoTo CloseAndExit
        End If
    Next longKishuCounter
    'ここまで来たという事は機種登録されてないという事
    boolRegistOK = False
'    MsgBox "機種登録されていないようなので、登録画面に移ります"
    strRegistRireki = strargRireki
    Call frmRegistNewKishu.Show
    '登録したので、もう1回リスト取得しに行く
    If boolRegistOK Then
        GoTo Serch_From_GlobalKishuList
    Else
        '機種登録OKフラグが立ってなかったら終了する
        Debug.Print "機種登録フラグNGにより終了"
        Exit Function
    End If
    Exit Function
CloseAndExit:
    Exit Function
ErrorCatch:
'    MsgBox "機種情報取得中にエラーが発生したようです"
    Debug.Print "getKishuInfoByRireki code: " & Err.Number & " Description: " & Err.Description
End Function
Public Function GetAllKishuInfo_Array() As typKishuInfo()
    '全機種情報をKishuInfo型の配列にして返す
    'ぐろばんる変数で共有しちゃおう？
    Dim arrKishuInfo() As typKishuInfo
    Dim isCollect As Boolean
    Dim dbKishuAll As clsSQLiteHandle
    Set dbKishuAll = New clsSQLiteHandle
    Dim intCounterKishu As Integer
    Dim varKishuTable As Variant
    Dim strSQLlocal As String
    On Error GoTo ErrorCatch
    '機種テーブルの有無を確認する
    If Not IsTableExist(Table_Kishu) Then
        MsgBox "機種テーブル（初期テーブル）が見つからなかったので新規作成します。"
        isCollect = InitialDBCreate
        If Not isCollect Then
            MsgBox "機種テーブルの作成に失敗したようです"
            GetAllKishuInfo_Array = arrKishuInfo
            GoTo CloseAndExit
            Exit Function
        End If
    End If
    'SQL作成
    strSQLlocal = "SELECT " & Kishu_Header & "," & Kishu_KishuName & "," & Kishu_KishuNickname & "," & _
                    Kishu_TotalKeta & "," & Kishu_RenbanKetasuu & _
                    " FROM " & Table_Kishu
    '機種テーブルの内容を配列で受け取る
    dbKishuAll.SQL = strSQLlocal
    isCollect = dbKishuAll.DoSQL_No_Transaction()
    If Not isCollect Then
'        MsgBox "SQL実行時に失敗したもよう"
        GoTo CloseAndExit
    End If
    If dbKishuAll.RecordCount = 0 Then
        Debug.Print "NoDataAvilable in T_Kishu"
        boolNoTableKishuRecord = True
        Exit Function
    End If
    '結果を配列で受け取る
    ReDim varKishuTable(dbKishuAll.RecordCount - 1)
    ReDim arrKishuInfo(dbKishuAll.RecordCount - 1)
    ReDim arrKishuInfoGlobal(UBound(arrKishuInfo))
    varKishuTable = dbKishuAll.RS_Array(boolPlusTytle:=False)
    Set dbKishuAll = Nothing
    'KishuInfo型に突っ込んでやる
    For intCounterKishu = LBound(varKishuTable, 1) To UBound(varKishuTable, 1)
        arrKishuInfo(intCounterKishu).KishuHeader = varKishuTable(intCounterKishu, 0)
        arrKishuInfo(intCounterKishu).KishuName = varKishuTable(intCounterKishu, 1)
        arrKishuInfo(intCounterKishu).KishuNickName = varKishuTable(intCounterKishu, 2)
        arrKishuInfo(intCounterKishu).TotalRirekiketa = varKishuTable(intCounterKishu, 3)
        arrKishuInfo(intCounterKishu).RenbanKetasuu = varKishuTable(intCounterKishu, 4)
    Next intCounterKishu
    arrKishuInfoGlobal = arrKishuInfo
    GetAllKishuInfo_Array = arrKishuInfo
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
    Set dbKishuAll = Nothing
    GetAllKishuInfo_Array = arrKishuInfo
    Exit Function
ErrorCatch:
    Debug.Print "GetAllKishu_Array code: " & Err.Number & "Description: " & Err.Description
End Function
Public Function GetFieldTypeNameByTableName(ByVal strargTableName As String) As Dictionary
    'テーブル名からフィールド名とデータタイプの一覧を取得する
    Dim dbFieldName As clsSQLiteHandle
    Set dbFieldName = New clsSQLiteHandle
    Dim isCollect As Boolean
    Dim dicFieldType As Dictionary
    Set dicFieldType = New Dictionary
    Dim varFieldType As Variant
    Dim intFieldCounter As Integer
    If Not IsTableExist(strargTableName) Then
        MsgBox strargTableName & " テーブルが見つかりませんでした。タイプ取得を中止します。"
        Set GetFieldTypeNameByTableName = dicFieldType
        GoTo CloseAndExit
    End If
    'マスターテーブルよりフィールド名とタイプ名を取得
    isCollect = dbFieldName.DoSQL_No_Transaction("SELECT name,type FROM pragma_table_info(""" & strargTableName & """)")
    varFieldType = dbFieldName.RS_Array(boolPlusTytle:=False)
    For intFieldCounter = LBound(varFieldType, 1) To UBound(varFieldType, 1)
        dicFieldType.Add varFieldType(intFieldCounter, 0), varFieldType(intFieldCounter, 1)
    Next intFieldCounter
    Set GetFieldTypeNameByTableName = dicFieldType
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
    Set dbFieldName = Nothing
    Set dicFieldType = Nothing
    Exit Function
End Function
 Public Function IsTableExist(ByVal strargTableName As String) As Boolean
    Dim dbExist As clsSQLiteHandle
    Set dbExist = New clsSQLiteHandle
    Dim isCollect As Boolean
    Dim strSQLlocal As String
    strSQLlocal = "SELECT tbl_name FROM sqlite_master WHERE type=""table"" AND name=""" & strargTableName & """"
    isCollect = dbExist.DoSQL_No_Transaction(strSQLlocal)
    If dbExist.RecordCount = 0 Then
        '検索結果にないので存在しない
        IsTableExist = False
    Else
        'テーブル発見
        IsTableExist = True
    End If
    Set dbExist = Nothing
    Exit Function
 End Function
Public Function GetRecordCountSimple(ByVal strargTableName As String, ByVal strargFieldName As String, Optional ByVal strargFindStr) As Long
    'テーブル名とフィールド名（一つ限定）、検索文字（省略可）を与えて、レコード数のみを返すシンプルなメソッド
    '検索文字列を与えない場合はcount()の簡易版として使えるかも
    '検索は WHERE (Field) (検索文字列)として行っています
    Dim dbSimple As clsSQLiteHandle
    Dim varReturnValue As Variant
    On Error GoTo ErrorCatch
    If Not IsTableExist(strargTableName) Then
        MsgBox strargTableName & "テーブルが見つかりません"
        GetRecordCountSimple = 0
        Exit Function
    End If
    Set dbSimple = New clsSQLiteHandle
    If strargFindStr = "" Then
        '検索文字列がない場合は素直にフィールド情報全部を対象に
        dbSimple.SQL = "SELECT COUNT(" & strargFieldName & ") FROM " & strargTableName
    Else
        '検索文字列がある場合は、Whereの条件に使ってやる
        dbSimple.SQL = "SELECT COUNT(" & strargFieldName & ") FROM " & strargTableName & _
        " WHERE " & strargFieldName & " " & strargFindStr
    End If
    dbSimple.DoSQL_No_Transaction
'    GetRecordCountSimple = dbSimple.RecordCount
    varReturnValue = dbSimple.RS_Array
    GetRecordCountSimple = varReturnValue(0, 0)
    Set dbSimple = Nothing
    Exit Function
ErrorCatch:
    GetRecordCountSimple = 0
    GetRecordCountSimple = errcOthers
    Set dbSimple = Nothing
    Debug.Print "SimpleRecorde code: " & Err.Number & "Description: " & Err.Description
    Exit Function
End Function
Public Function GetKishuinfoByZuban(strargZuban As String) As typKishuInfo
    '図番をもとにKishuInfoを引っ張ってくる
    Dim Kishu As typKishuInfo
    Dim longKishuCounter As Long
    On Error GoTo ErrorCatch
    If strargZuban = "" Then
        Debug.Print "GetKishuInfoByZXuban No arg"
        GoTo CloseAndExit
    End If
    'ぐろばんるのが初期化されているかチェック
    If (Not arrKishuInfoGlobal) = -1 Then
        'ここに来ると未初期化らしい・・・
        Call GetAllKishuInfo_Array
    End If
    If boolNoTableKishuRecord = True Then
        '機種テーブルが空の場合は終了する
        GoTo CloseAndExit
    End If
    '機種テーブルarryaを全部調べる
    For longKishuCounter = LBound(arrKishuInfoGlobal, 1) To UBound(arrKishuInfoGlobal, 1)
        If strargZuban = arrKishuInfoGlobal(longKishuCounter).KishuName Then
            '図番とKishuNameが一致したらKishuInfoを返してやる
            Kishu.KishuHeader = arrKishuInfoGlobal(longKishuCounter).KishuHeader
            Kishu.KishuName = arrKishuInfoGlobal(longKishuCounter).KishuName
            Kishu.KishuNickName = arrKishuInfoGlobal(longKishuCounter).KishuNickName
            Kishu.TotalRirekiketa = arrKishuInfoGlobal(longKishuCounter).TotalRirekiketa
            Kishu.RenbanKetasuu = arrKishuInfoGlobal(longKishuCounter).RenbanKetasuu
            GoTo CloseAndExit
        End If
    Next longKishuCounter
    '機種が見つからなかったので、そのまま終了
    GoTo CloseAndExit
CloseAndExit:
    GetKishuinfoByZuban = Kishu
    Exit Function
ErrorCatch:
'    MsgBox "機種情報取得中にエラーが発生したようです"
    Debug.Print "getKishuInfoByRireki code: " & Err.Number & " Description: " & Err.Description
End Function
Public Function GetLocalTimeWithMilliSec() As String
    '現在日時をミリ秒まで付けて、フォーマット済みStringとして返す
    'ISO1806形式
    'yyyy-mm-ddTHH:MM:SS.fff
    Dim strDateWithMillisec As String
    Dim timeLocalTime As SYSTEMTIME
    Call GetLocalTime(timeLocalTime)
    strDateWithMillisec = ""
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wYear, "0000")
    strDateWithMillisec = strDateWithMillisec & "-"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wMonth, "00")
    strDateWithMillisec = strDateWithMillisec & "-"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wDay, "00")
    strDateWithMillisec = strDateWithMillisec & "T"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wHour, "00")
    strDateWithMillisec = strDateWithMillisec & ":"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wMinute, "00")
    strDateWithMillisec = strDateWithMillisec & ":"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wSecond, "00")
    strDateWithMillisec = strDateWithMillisec & "."
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wMilliseconds, "000")
    GetLocalTimeWithMilliSec = strDateWithMillisec
End Function
Public Function GetLastRireki(ByVal strargTableName As String) As String
    '与えられたテーブル名の最後の履歴を取得する
    Dim dbLastRireki As clsSQLiteHandle
    Dim varResult As Variant
    On Error GoTo ErrorCatch
    If Not IsTableExist(strargTableName) Then
        Debug.Print "GetLastRireki Table: " & strargTableName & " not found"
        GetLastRireki = ""
        Exit Function
    End If
    Set dbLastRireki = New clsSQLiteHandle
    dbLastRireki.SQL = "SELECT " & Job_Rireki & " FROM (SELECT " & _
                        Job_Rireki & ",MAX(" & Job_RirekiNumber & ") FROM " & _
                        strargTableName & ");"
    Call dbLastRireki.DoSQL_No_Transaction
    varResult = dbLastRireki.RS_Array(boolPlusTytle:=False)
    GetLastRireki = CStr(varResult(0, 0))
    GoTo CloseAndExit
ErrorCatch:
    Debug.Print "GetLastRireki code : " & Err.Number & "Description: " & Err.Description
    GetLastRireki = ""
    GoTo CloseAndExit
CloseAndExit:
   Set dbLastRireki = Nothing
   Exit Function
End Function
Public Function GetNextRireki(ByVal strargTableName As String) As String
    '与えられたテーブルの次の履歴を取得する
    Dim strLastRireki As String
    Dim strNewRireki As String
    Dim KishuLocal As typKishuInfo
    'テーブルがない場合や、ラスト履歴が空白だったら空白返して終了
    If Not IsTableExist(strargTableName) Then
        Debug.Print "GetNextRireki Table: " & strargTableName & " not found"
        GetNextRireki = ""
        Exit Function
    End If
    strLastRireki = GetLastRireki(strargTableName)
    If strLastRireki = "" Then
        Debug.Print "GetNextRireki : Last Rireki Empty"
        GetNextRireki = ""
        Exit Function
    End If
    KishuLocal = getKishuInfoByRireki(strLastRireki)
    If KishuLocal.RenbanKetasuu = 0 Then
        Debug.Print "GetNextRireki : KishuInfo Empty"
        GetNextRireki = ""
        Exit Function
    End If
    strNewRireki = Mid(strLastRireki, 1, KishuLocal.TotalRirekiketa - KishuLocal.RenbanKetasuu) & _
                    Right(String$(KishuLocal.RenbanKetasuu, "0") & CStr((CLng(Right(strLastRireki, KishuLocal.RenbanKetasuu)) + 1)), KishuLocal.RenbanKetasuu)
    GetNextRireki = strNewRireki
    Exit Function
End Function
Public Sub OutputArrayToCSV(ByRef vararg2DimentionsDataArray As Variant, ByVal strargFilePath As String, Optional ByVal strargFileEncoding As String = "UTF-8")
    '二次元配列をCSVに吐き出す
    Dim byteDimentions As Byte
    Dim objFileStream As ADODB.Stream
    Dim longRowCounter As Long
    Dim longFieldCounter As Long
    Dim strarrField() As String
    Dim strLineBuffer As String
    On Error GoTo ErrorCatch
    byteDimentions = getArryDimmensions(vararg2DimentionsDataArray)
    If Not byteDimentions = 2 Then
        MsgBox "引数に二次元配列以外が与えられました。処理を中止します。"
        Debug.Print "OutputArrayToCSV : Not 2 Dimension Array"
        Exit Sub
    End If
    Set objFileStream = New ADODB.Stream
    With objFileStream
        'エンコード指定
        .Charset = strargFileEncoding
        '改行コード指定
        .LineSeparator = adCRLF
        .Open
        '行数ループ
        For longRowCounter = LBound(vararg2DimentionsDataArray, 1) To UBound(vararg2DimentionsDataArray, 1)
            'フィールド数ループ、ここでラインバッファを組み立てる
            'まずはstring配列にフィールド情報を入れて、Joinで連結する
            ReDim strarrField(UBound(vararg2DimentionsDataArray, 2))
            For longFieldCounter = LBound(vararg2DimentionsDataArray, 2) To UBound(vararg2DimentionsDataArray, 2)
                If IsNull(vararg2DimentionsDataArray(longRowCounter, longFieldCounter)) Then
                    'Nullの場合はNULLを入入力してやる
                    strarrField(longFieldCounter) = "NULL"
                Else
                    '通常はこっち
                    strarrField(longFieldCounter) = CStr(vararg2DimentionsDataArray(longRowCounter, longFieldCounter))
                End If
            Next longFieldCounter
            strLineBuffer = Join(strarrField, ",")
            .WriteText strLineBuffer, adWriteLine
        Next longRowCounter
        'ループが終わったらテキストファイル書き出す（上書き保存）
        .SaveToFile strargFilePath, adSaveCreateOverWrite
        .Close
    End With
    MsgBox "CSV出力完了 " & strargFilePath
    Exit Sub
ErrorCatch:
    Debug.Print "OutputArrayToCSV code: " & Err.Number & " Description: " & Err.Description
    Exit Sub
End Sub