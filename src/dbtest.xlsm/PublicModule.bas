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
Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
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
Public Function InitialDBCreate() As Boolean
    Dim isCollect As Boolean
    Dim strSQL  As String
    Dim dbSQLite3 As clsSQLiteHandle
    Set dbSQLite3 = New clsSQLiteHandle
    On Error GoTo ErrorCatch
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
    'カレントディレクトリをブックのディレクトリに変更
    ChCurrentDirW (ThisWorkbook.Path)
    strCurrentDir = CurDir
    'データベースディレクトリはある前提で進めるので不要
    'そこにSQLiteのDLLあるから！
'    'DataBaseディレクトリの存在有無確認"
'    If fso.FolderExists(constDatabasePath) <> True Then
'        'ディレクトリ存在しない場合作成しよ？
'        MsgBox "データベースフォルダが無いため作成します。"
'        MkDir constDatabasePath
'    End If
    'データベースディレクトリに移動
    strCurrentDir = CurDir & "\" & constDatabasePath
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
    strSQL = strSQL & """" & BarcordNumber & """ TEXT NOT NULL," & vbCrLf
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
    strSQL = strSQL & """" & BarcordNumber & """ TEXT NOT NULL," & vbCrLf
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
    CreateTable_by_KishuName = True
    Exit Function
End Function
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