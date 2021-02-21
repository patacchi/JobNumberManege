VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLiteHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private myDbHandle As LongPtr           'DBハンドル
Private RetVal As Long                  '主にSQLite3用の帰り値
Private myStmtHandle As LongPtr         'SQLステートメントハンドル
Private strDBPath As String             'DBファイルパス名
Private longDBRowIndex As Long          '現在のデータベールカーソル行インデックス
Private strSQL As String                'SQL文字列
Private myRecord As Variant             '結果として帰す配列（バリアント）myRecord(Array(),....)、先頭1行目にフィールド名入れる事にした
Private isInitialize As Boolean         'イニシャライズ済みかどうか
Private isDBOpen As Boolean             'DBがオープン済みかどうか
Private isStatement As Boolean          'Statementのハンドルが有効（True）かどうか
Private longRecordCount As Long         'Select時の （SQLITE_ROWの）数 = ヒットしたレコードの数を記録
Private Sub Class_Initialize()
    'コンストラクタ（クラスインスタンス生成時に最初に実行される）
    isInitialize = False
    isDBOpen = False
    longDBRowIndex = 0
    longRecordCount = 0
    myRecord = Array(0)
End Sub
Private Sub Class_Terminate()
    'デストラクタ（クラス破棄する時に実行される）
    Set myRecord = Nothing
    strSQL = Empty
    '基本的にお掃除は各自でやるので、まだデータベース接続されてる時だけにやる
    If isStatement Then
        'ステートメントハンドルお掃除
        RetVal = SQLite3Finalize(myStmtHandle)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLiteFinalize returned:" & RetVal, SQLite3ErrMsg(myStmtHandle)
        End If
        isStatement = False
    End If
    If isDBOpen Then
        'DB接続クローズ
        RetVal = SQLite3Close(myDbHandle)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Close returned:" & RetVal, SQLite3ErrMsg(myStmtHandle)
        End If
        isDBOpen = False
    End If
    Exit Sub
End Sub
Property Get RS() As Variant
    'レコードセット取得（一行）
    If longDBRowIndex > Me.RecordCount Then
        MsgBox "レコードセット取得時にエラー発生。レコード総数以上のインデックスが指定されました"
        RS = Nothing
        Exit Property
    End If
    RS = myRecord(longDBRowIndex)
End Property
Property Get RS_Array_Dictionary() As Dictionary
    'RecoredSetをそのまま配列で返せないかと
    '扱いやすいように二次元配列に入れなおしてから返す
    '更にDictionary型で返すようにした
    Dim dicArrayLocal As Dictionary
    Dim strArrFieldName() As String
    Dim longRowCounter As Long
    Dim longColumnCount As Long
    Set dicArrayLocal = New Dictionary
    On Error GoTo ErrorCatch
    '1次元目の要素数分ループ
    'フィールドのリストを取得
    ReDim strArrFieldName(UBound(myRecord(LBound(myRecord, 1)), 1))
    For longColumnCount = LBound(myRecord(LBound(myRecord, 1)), 1) To UBound(myRecord(LBound(myRecord, 1)), 1)
        strArrFieldName(longColumnCount) = myRecord(LBound(myRecord, 1))(longColumnCount)
    Next longColumnCount
    'タイトル行を取得しない場合はこっち
    '行数分ループ
    For longRowCounter = LBound(myRecord, 1) + 1 To UBound(myRecord, 1)
        dicArrayLocal.Add CStr((longRowCounter)), New Dictionary
        For longColumnCount = LBound(myRecord(longRowCounter), 1) To UBound(myRecord(longRowCounter), 1)
'            varArray(longRowCounter - 1, longColumnCount) = myRecord(longRowCounter)(longColumnCount)
            dicArrayLocal(CStr(longRowCounter)).Add (strArrFieldName(longColumnCount)), myRecord(longRowCounter)(longColumnCount)
        Next longColumnCount
    Next longRowCounter
    Set RS_Array_Dictionary = dicArrayLocal
    Exit Property
ErrorCatch:
    Debug.Print "RS_Array code: " & Err.Number & " Desctiption: " & Err.Description
End Property
Property Get RS_Array(Optional ByRef boolPlusTytle As Boolean) As Variant
    'RecoredSetをそのまま配列で返せないかと
    '扱いやすいように二次元配列に入れなおしてから返す
    Dim varArray As Variant
    Dim longRowCounter As Long
    Dim longColumnCount As Long
    On Error GoTo ErrorCatch
    varArray = Array()
    '1次元目の要素数分ループ
    If boolPlusTytle = True Then
        ReDim varArray(UBound(myRecord, 1), UBound(myRecord(0), 1))
        For longRowCounter = LBound(myRecord, 1) To UBound(myRecord, 1)
            For longColumnCount = LBound(myRecord(longRowCounter), 1) To UBound(myRecord(longRowCounter), 1)
                varArray(longRowCounter, longColumnCount) = myRecord(longRowCounter)(longColumnCount)
            Next longColumnCount
        Next longRowCounter
    Else
        'タイトル行を取得しない場合はこっち
        ReDim varArray(UBound(myRecord, 1) - 1, UBound(myRecord(0), 1))
        For longRowCounter = LBound(myRecord, 1) + 1 To UBound(myRecord, 1)
            For longColumnCount = LBound(myRecord(longRowCounter), 1) To UBound(myRecord(longRowCounter), 1)
                varArray(longRowCounter - 1, longColumnCount) = myRecord(longRowCounter)(longColumnCount)
            Next longColumnCount
        Next longRowCounter
    End If
    RS_Array = varArray
    Exit Property
ErrorCatch:
    Debug.Print "RS_Array code: " & Err.Number & " Desctiption: " & Err.Description
End Property
Property Get RecordCount() As Long
    'レコード総数取得
    On Error GoTo ErrorCatch
'    RecordCount = UBound(myRecord) - LBound(myRecord) + 1
    '先頭行にタイトルを入れる事にしたので行数調整
'    RecordCount = UBound(myRecord) - LBound(myRecord)
    RecordCount = longRecordCount
    Exit Property
ErrorCatch:
    RecordCount = 0
    Debug.Print "Error Ocure SQLiteHande Class. code:" & Err.Number & " Description:" & Err.Description
End Property
Property Get EOF() As Boolean
    'データベースがEOFかどうか調べて返す
    If longDBRowIndex = Me.RecordCount() Then
        'レコードカウントとDBカーソルインデックスが同じなので最後
        EOF = True
    Else
        EOF = False
    End If
End Property
Property Let SQL(ByVal strargSQL As String)
    'SQL文設定
    strSQL = strargSQL
End Property
Property Get SQL() As String
    'SQL文取得
    SQL = strSQL
End Property
Public Sub MoveNext(Optional ByVal longDBRow As Long)
    'DBレコードセットのステートメントカーソルを次の行へ、インデックス番号が指定されている場合はそこから
    If longDBRow > 0 Then
        longDBRowIndex = longDBRow - 1
    End If
    longDBRowIndex = longDBRowIndex + 1
End Sub
Public Sub MoveFirst()
    'DBステートメントカーソルを最初に戻す
    longDBRowIndex = 0
End Sub
Public Function Initialize(Optional ByRef strLibDirarg As String) As Boolean
    Dim isCollect As Boolean
    'イニシャライズ動作、まだDBに接続はしない
    On Error GoTo ErrorCatch
    '特に指定されていなければ標準のDBファイルとディレクトリを指定する
    If strLibDirarg = "" Then
        'DBイニシャライズ
        isCollect = DBInishialize()
    Else
        'DBイニシャライズ（ライブラリディレクトリ指定）
        isCollect = DBInishialize(strLibDirarg)
    End If
    If isCollect Then
        Initialize = True
        isInitialize = True
        Exit Function
    Else
        '初期化失敗
        MsgBox "DB初期化失敗"
        Initialize = False
        Exit Function
    End If
ErrorCatch:
    Initialize = False
    Debug.Print "Error in Initialize. code:" & Err.Number & " Description:" & Err.Description
    Class_Terminate
End Function
Public Function DoSQL_No_Transaction(Optional ByVal strSQLarg As String, Optional strDBPath As String) As Boolean
    'SQL文字列とDBパス名（オプション）から、トランザクションなしの処理を行います（軽いの専用）
    Dim stepMsg As String
    Dim isCollect As Boolean
    Dim strErrorMsg As String
    On Error GoTo ErrorCatch
    'まだDBイニシャライズしてないようならここでやる
    If Not isInitialize Then
        'DBイニシャライズ
        isCollect = DBInishialize()
    End If
    'DB接続の有無を確認し、接続してなかったらここで接続する
    If Not isDBOpen Then
        'DBファイル名
        If strDBPath = "" Then
            strDBPath = CurDir & "\" & constJobNumberDBname
        End If
        'Job情報DBファイルオープン（無ければ作るはず）
        RetVal = SQLite3Open(strDBPath, myDbHandle)
        If RetVal <> SQLITE_OK Then
            strErrorMsg = "SQLite3Open returned" & RetVal & vbTab & SQLite3ErrMsg(myDbHandle)
            'Debug.Print "SQLite3Open returned" & RetVal, SQLite3ErrMsg(myDbHandle)
'            Debug.Print strErrorMsg
'            myRecord = Array()
'            ReDim myRecord(0)
'            myRecord(0) = Array(strErrorMsg)
            DoSQL_No_Transaction = False
            GoTo NoTitleExit
            Exit Function
        End If
        'DB接続フラグTrue
        isDBOpen = True
    End If
    'SQL文字列コンパイル
    'SQL指定されていればその内容で実行する
    '事前にプロパティで設定してもいいと思う
    If Not strSQLarg = "" Then
        Me.SQL = strSQLarg
    End If
    RetVal = SQLite3PrepareV2(myDbHandle, Me.SQL, myStmtHandle)
    isStatement = True
    If RetVal <> SQLITE_OK Then
        strErrorMsg = "SQLite3PrepareV2 returned" & RetVal & vbTab & SQLite3ErrMsg(myDbHandle)
        'Debug.Print "SQLite3PrepareV2 returned" & RetVal, SQLite3ErrMsg(myDbHandle)
        DoSQL_No_Transaction = False
        GoTo NoTitleExit
    End If
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_OK Or RetVal = SQLITE_DONE Then
        'ここで正常終了
        'Insertとか結果が一つのはこっちになるはず
        'Sqlite3Changesで結果行数取得してそれで終わり
        strErrorMsg = "直前の操作で " & SQLite3Changes(myDbHandle) & " 個の変更がありました（正常終了）"
'        Debug.Print strErrorMsg
'        myRecord = Array()
'        ReDim myRecord(0)
'        myRecord(0) = Array(strErrorMsg)
        DoSQL_No_Transaction = True
'        GoTo CloseAndExit
        GoTo NoTitleExit
        Exit Function
    ElseIf RetVal = SQLITE_ROW Then
        'SQL実行とデータ配列格納をここからやる
        isCollect = Get_Data_and_Put_Arry
    Else
        'ここはもう異常終了でいいと思う
        DoSQL_No_Transaction = False
        GoTo CloseAndExit
    End If
    If isCollect Then
        '正常終了したもより
        DoSQL_No_Transaction = True
        GoTo CloseAndExit
        Exit Function
    Else
        '何かあったっぽい
        DoSQL_No_Transaction = False
        GoTo CloseAndExit
    End If
    '正常終了
    DoSQL_No_Transaction = True
    GoTo CloseAndExit
    Exit Function
NoTitleExit:
        Debug.Print strErrorMsg
        myRecord = Array()
        ReDim myRecord(1)
        myRecord(0) = Array("No Title")
        myRecord(1) = Array(strErrorMsg)
        GoTo CloseAndExit
ErrorCatch:
    If Err.Number <> 0 Then
'        MsgBox Err.Number & vbCrLf & Err.Description
    End If
    Debug.Print "DoSQL_No_Transaction code: " & Err.Number & "description: " & Err.Description
    DoSQL_No_Transaction = False
    Exit Function
CloseAndExit:
    'StatamentHandle お掃除
    If isStatement Then
        RetVal = SQLite3Finalize(myStmtHandle)
        If RetVal <> SQLITE_OK Then
            strErrorMsg = "SQLite3Finalize returned" & RetVal & vbTab & SQLite3ErrMsg(myDbHandle)
            'Debug.Print "SQLite3Finalize returned" & RetVal, SQLite3ErrMsg(myDbHandle)
            Debug.Print strErrorMsg
            myRecord = Array()
            ReDim myRecord(0)
            myRecord(0) = Array(strErrorMsg)
            DoSQL_No_Transaction = False
        End If
        isStatement = False
    End If
    'DB接続クローズ
    'DB接続閉じるのはクラス破棄されるまで待っててもいいかな？
'    RetVal = SQLite3Close(myDbHandle)
'    If RetVal <> SQLITE_OK Then
'        Debug.Print "SQLite3Close returned" & RetVal, SQLite3ErrMsg(myDbHandle)
'        Exit Function
'    End If
'    'DB接続フラグFalseに
'    isDBOpen = False
    Exit Function
End Function
Public Function Do_SQL_With_Transaction(Optional ByVal strSQLarg As String, Optional strDBPath As String) As Boolean
    'トランザクション処理ありのSQL実行
    Dim sqlite3WithTrunsaction As clsSQLiteHandle
    Set sqlite3WithTrunsaction = New clsSQLiteHandle
    Dim isCollect As Boolean
    Dim strOrigSQL As String
    On Error GoTo RollBackAndExit
    strOrigSQL = Me.SQL
    'トランザクション開始
    isCollect = DoSQL_No_Transaction("BEGIN TRANSACTION")
    If Not isCollect Then
        GoTo RollBackAndExit
    End If
    Me.SQL = strOrigSQL
    isCollect = DoSQL_No_Transaction()
    If Not isCollect Then
        GoTo RollBackAndExit
    End If
    Do_SQL_With_Transaction = True
    '成功したのでコミットを
    isCollect = DoSQL_No_Transaction("COMMIT TRANSACTION")
    If Not isCollect Then
        GoTo RollBackAndExit
    End If
    Exit Function
RollBackAndExit:
    'トランザクションをロールバックしたうえで、終了する
    DoSQL_No_Transaction ("ROLLBACK TRANSACTION")
    MsgBox "ロールバック発生"
    Do_SQL_With_Transaction = False
    Debug.Print "Do_SQL_With_Transaction code: " & Err.Number & "Description: " & Err.Description
    Set sqlite3WithTrunsaction = Nothing
    Exit Function
End Function
Private Function Get_Data_and_Put_Arry() As Boolean
    'SQLiteステートメントに対し、SQLITE_DONEが出てくるまで配列にデータを突っ込む
    '次のデータがある場合はSQLITE_ROWが帰ってきてるはず
    'SQLITE_DONE、SQL_OK以外で終わった場合はFalseを返す
    Dim longColCount As Long            'フィールド総数
    Dim longCounterRow As Long
    Dim longCounterColumn As Long
    Dim longColType As Long             'Columnの型（ColumnValueの時に使う）
    Dim strColName As String            'フィールド名
    Dim varColumnData As Variant        'DBの各フィールドの値を格納するArray()、最初の行はフィールド名入ってるよ
    Dim dblStartTime As Double
    'SQL実行
    On Error GoTo ErrorCatch
    If (RetVal <> SQLITE_OK) And (RetVal <> SQLITE_DONE) And (RetVal <> SQLITE_ROW) Then
        '異常値なので即終了
        Debug.Print "SQLite3Step returned" & RetVal, SQLite3ErrMsg(myDbHandle)
        Get_Data_and_Put_Arry = False
        Exit Function
    End If
    'とりあえず総行数を確認する方法が分からなかったので、固定値で配列確保(ToDo)
    myRecord = Array()
    ReDim myRecord(constDefaultArraySize)
    varColumnData = Array()
    longCounterRow = 0
    'フィールド総数取得
    longColCount = SQLite3ColumnCount(myStmtHandle)
    '列データの配列サイズ調整
    ReDim varColumnData(longColCount - 1)
    dblStartTime = timer()
    'データの1行目にフィールドタイトル名を入れる
    'RetValueの結果によって処理を分岐（複数あるのか、1行だけなのか
    Select Case RetVal
    Case SQLITE_OK Or SQLITE_DONE
        '多分結果1行のみここっち
        myRecord(longCounterRow) = "No Field Title"
        longCounterRow = longCounterRow + 1
    Case SQLITE_ROW
        'フィールド数分ループ
        For longCounterColumn = 0 To longColCount - 1
            longColType = SQLite3ColumnType(myStmtHandle, longCounterColumn)
            varColumnData(longCounterColumn) = SQLite3ColumnName(myStmtHandle, longCounterColumn)
        Next longCounterColumn
        myRecord(longCounterRow) = varColumnData
'        varColumnData = Array()
        longCounterRow = longCounterRow + 1
'        longRecordCount = longRecordCount + 1
    Case Else
        'それ以外？そんははずは・・・
        Debug.Print "前提のIfの漏れです、何とかして・・・"
        longCounterRow = longCounterRow + 1
    End Select
    Do While RetVal = SQLITE_ROW
        '配列に突っ込む処理開始
        'フィールドの数分ループする
        For longCounterColumn = 0 To longColCount - 1
            'データタイプ取得し、値を取得する
            longColType = SQLite3ColumnType(myStmtHandle, longCounterColumn)
            varColumnData(longCounterColumn) = ColumnValue_InClass(myStmtHandle, longCounterColumn, longColType)
        Next longCounterColumn
        myRecord(longCounterRow) = varColumnData
'        varColumnData = Array()
        ' Rowカウンターインクリメント
        'レコードカウンターもインクリメントしてあげて・・・
        longCounterRow = longCounterRow + 1
        longRecordCount = longRecordCount + 1
        If UBound(myRecord) - longCounterRow <= 50 Then
            '初期確保上限まで50切ったら増量分だけ増やしてやる（重いよ・・・
            ReDim Preserve myRecord(UBound(myRecord) + constAddArraySize)
        End If
        RetVal = SQLite3Step(myStmtHandle)
    Loop
    'ダウンサイズ
    ReDim Preserve myRecord(CLng(longCounterRow - 1))
    Debug.Print "SQL→配列格納所要時間:" & timer() - dblStartTime
    '正常終了
    Get_Data_and_Put_Arry = True
    Exit Function
ErrorCatch:
    Get_Data_and_Put_Arry = False
    MsgBox "SQL結果配列書き出し中にエラー"
    Debug.Print "Get_Data_and_Put_Array code:" & Err.Number & " Description:" & Err.Description
    Exit Function
End Function
#If Win64 Then
Private Function ColumnValue_InClass(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#Else
Private Function ColumnValue_InClass(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#End If
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue_InClass = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue_InClass = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue_InClass = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue_InClass = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue_InClass = Null
    End Select
End Function
Private Function DBInishialize(Optional ByRef strLibDir As String) As Boolean
    If strLibDir = "" Then
        'Libディレクトリ指定されなかった場合は、DBディレクトリと同一とする（お引越し楽だから・・・）
        'strLibDir = ThisWorkbook.Path + "\x64"
'        strLibDir = ThisWorkbook.Path & "\" & constDatabasePath
        strLibDir = constDatabasePath
    End If
    'イニシャライズ
    Dim InitReturn As Long
    Dim isCollect
    #If Win64 Then
        ' I put the 64-bit version of SQLite.dll under a subdirectory called x64
        InitReturn = SQLite3Initialize(strLibDir)
    #Else
        InitReturn = SQLite3Initialize ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    #End If
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        DBInishialize = False
        Exit Function
    End If
    'DBディレクトリに移動
'    isCollect = ChcurrentforDB()
    isCollect = ChCurrentToDBDirectory
    If Not isCollect Then
        MsgBox "DBディレクトリ存在確認中にエラー発生"
        DBInishialize = False
        Exit Function
    End If
    DBInishialize = True
    isInitialize = True
End Function