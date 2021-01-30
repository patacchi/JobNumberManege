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


Public Function ChCurrentDirW(ByVal DirName As String)
    'UNICODE対応ChCurrentDir
    'SetCurrentDirectoryW（UNICODE）なので
    'StrPtrを介す必要がある・・？
    SetCurrentDirectoryW StrPtr(DirName)
End Function

Public Function IsDBFileExist() As Boolean
    'DBファイルの存在を確認し、無い場合はDBファイル二つ作成し（中身空）
    '既にカレントディレクトリがDBディレクトリにある前提で動いてます
    'ジョブ情報DBにのみ、空のT_Kishu、機種情報格納テーブルを設定する
    On Error GoTo ErrorCatch
    Dim fso As New Scripting.FileSystemObject

    'ジョブ情報DB有無チェック
    'ジョブ情報DBが存在しない場合は、DB存在無しとして扱う
    If fso.FileExists(constJobNumberDBname) <> True Then
        MsgBox ("DBファイルが見つからなかったため新規作成します")
        'DB新規作成処理
        If Not InitialDBCreate Then
            MsgBox ("DBファイルチェック（新規DB作成）でエラー")
            IsDBFileExist = False
            Set fso = Nothing
            Exit Function
        End If
    End If
    IsDBFileExist = True
    Set fso = Nothing
    Exit Function

ErrorCatch:
'    MsgBox ("DBファイル存在確認中にエラー発生")
    Debug.Print "IsDBFileExist Error code:" & Err.Number & "Description: " & Err.Description
    Set fso = Nothing
    Exit Function
End Function

Public Function InitialDBCreate() As Boolean
    Dim isCollect As Boolean
    Dim strSQL  As String
    Dim dbSqlite3 As clsSQLiteHandle
    Set dbSqlite3 = New clsSQLiteHandle
    
    On Error GoTo ErrorCatch
    '初期テーブル作成用SQL文作成(T_Kishu)
    strSQL = ""
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_Kishu & """("""
    strSQL = strSQL & Kishu_Header & """ TEXT NOT NULL UNIQUE,"""
    strSQL = strSQL & Kishu_KishuName & """ TEXT NOT NULL UNIQUE,"""
    strSQL = strSQL & Kishu_KishuNickname & """ TEXT NOT NULL ,"""
    strSQL = strSQL & Kishu_TotalKeta & """ NUMERIC NOT NULL,"""
    strSQL = strSQL & Kishu_RenbanKetasuu & """ NUMERIC NOT NULL,"""
    strSQL = strSQL & Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP,"""
    strSQL = strSQL & Field_Update & """ TEXT)"
    
    isCollect = dbSqlite3.DoSQL_No_Transaction(strSQL)
    Set dbSqlite3 = Nothing
    'テスト実装_SQL作成テスト
    isCollect = CreateTable_by_KishuName("Test15")
    If Not isCollect Then
        InitialDBCreate = False
    End If
    '正常終了
    InitialDBCreate = True
    Exit Function
ErrorCatch:
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
    End If
    Exit Function
End Function
Public Function registNewKishu(ByVal strKishuheader As String, ByVal strKishuname As String, ByVal strKishuNickname As String, _
                                ByVal byteTotalKetasu As Byte, ByVal byteRenbanKetasu As Byte) As Boolean
    '新機種登録動作
End Function
Public Function ChcurrentAndReturnDBName()
    'Activesheetの設定で、カレントディレクトリを移動し
    '更にデータベース名を返す(String)
    'カレントディレクトリの取得（UNCパス対応）
    Dim strCurrentDir As String
    Dim fso As New Scripting.FileSystemObject
    Dim rngName As Name
    Dim strDatabaseName As String
    
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
End Function

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
    Dim dbSqlite3 As clsSQLiteHandle
    Set dbSqlite3 = New clsSQLiteHandle
    
    'T_Jobdata ジョブの履歴とJob番号
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_JobDataPri & strKishuname & """ (""" & _
    Job_Number & """ TEXT NOT NULL,""" & _
    Job_RirekiHeader & """ TEXT NOT NULL,""" & _
    Job_RirekiNumber & """ NUMERIC NOT NULL UNIQUE,""" & _
    Job_Rireki & """ TEXT NOT NULL UNIQUE,""" & _
    Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP,""" & _
    Field_Update & """ TEXT," & _
    "Primary Key(""" & Job_Rireki & """)" & _
    ");"
    isCollect = dbSqlite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    strSQL = ""
    
    '続いてT_Barcorde バーコードテーブル（ピッてするやつ）
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_Barcodepri & strKishuname & """ (""" & _
    Job_Number & """ TEXT,""" & _
    BarcordNumber & """ TEXT NOT NULL,""" & _
    Laser_Rireki & """ TEXT NOT NULL UNIQUE,""" & _
    Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP,""" & _
    Field_Update & """ TEXT," & _
    "Primary Key(""" & Laser_Rireki & """)" & _
    ");"
    isCollect = dbSqlite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    strSQL = ""

    '最後にリトライ履歴（いるの？これ）
    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_Retrypri & strKishuname & """ (""" & _
    Job_Number & """ TEXT,""" & _
    BarcordNumber & """ TEXT NOT NULL,""" & _
    Laser_Rireki & """ TEXT NOT NULL,""" & _
    Retry_Reason & """ TEXT,""" & _
    Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP,""" & _
    Field_Update & """ TEXT" & _
    ");"
    isCollect = dbSqlite3.DoSQL_No_Transaction(strSQL)
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
    isCollect = dbSqlite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    'ジョブ履歴テーブル
    strSQL = ""
    strSQL = "CREATE UNIQUE INDEX IF NOT EXISTS ""ix" & Table_JobDataPri & strKishuname & """ ON """ & _
    Table_JobDataPri & strKishuname & """ (""" & Job_Rireki & """ ASC);"
    isCollect = dbSqlite3.DoSQL_No_Transaction(strSQL)
    If Not isCollect Then
        CreateTable_by_KishuName = False
        Exit Function
    End If
    'インデックス作成終了
    Set dbSqlite3 = Nothing
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
    Dim Kishu As typKishuInfo
    Dim dbSqlite3 As clsSQLiteHandle
    dbSqlite3 = New clsSQLiteHandle
    Dim strSQLlocal As String
    Dim strarrKishuHeader() As String
    Dim longKishusuCounter As Long

    If strargRireki = "" Then
        MsgBox "機種情報検索には履歴が必須です"
        getKishuInfoByRireki = Kishu
        Exit Function
    End If

    '機種ヘッダのみのリストを受け取る
    strSQLlocal = "SELECT " & Kishu_Header & _
                    "FROM " & Table_Kishu
    
    
    
End Function

 
