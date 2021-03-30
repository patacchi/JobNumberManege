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
Private boolLocalMode As Boolean        'ローカルDB使用モード(openの時に条件分岐する)
Private strDBPath As String             'DBファイルパス名
Private longDBRowIndex As Long          '現在のデータベールカーソル行インデックス
Private strSQL As String                'SQL文字列
Private myRecord As Variant             '結果として帰す配列（バリアント）myRecord(Array(),....)、先頭1行目にフィールド名入れる事にした
Private isInitialize As Boolean         'イニシャライズ済みかどうか
Private isDBOpen As Boolean             'DBがオープン済みかどうか
Private isStatement As Boolean          'Statementのハンドルが有効（True）かどうか
Private longRecordCount As Long         'Select時の （SQLITE_ROWの）数 = ヒットしたレコードの数を記録
Private dicNamedParm As Dictionary      'keyにNamedパラメータ名、Valueに(Type=BindType.???,Data=実際のデータ)というDictionaryを格納
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
    Set dicNamedParm = Nothing
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
Property Let DBPath(ByVal strargDBPath As String)
    'データベースパス設定
    strDBPath = strargDBPath
End Property
Property Get DBPath() As String
    'データベースパス取得
    DBPath = strDBPath
End Property
Property Let DbHandle(ByVal longptrDBHandle As LongPtr)
    'DBハンドルを設定
    myDbHandle = longptrDBHandle
End Property
Property Get DbHandle() As LongPtr
    'DBハンドルを取得
    DbHandle = myDbHandle
End Property
Property Let StmtHandle(ByVal longptrStmtHandle As LongPtr)
    'SQLステートメントのハンドルを設定
    myStmtHandle = longptrStmtHandle
End Property
Property Get StmtHandle() As LongPtr
    'SQLステートメントハンドルの取得
    StmtHandle = myStmtHandle
End Property
Property Let LocalMode(ByVal boolargLocal As Boolean)
    'ローカルモードフラグ設定
    boolLocalMode = boolargLocal
End Property
Property Get LocalMode() As Boolean
    'ローカルモード取得
    LocalMode = boolLocalMode
End Property
Property Set NamedParm(dicargParmList As Dictionary)
    If dicNamedParm Is Nothing Then
        Set dicNamedParm = New Dictionary
    End If
    If dicargParmList Is Nothing Then
        Exit Property
    End If
    dicNamedParm.Add dicargParmList.Keys(0), dicargParmList.Items(0)
'    Set dicNamedParm = dicargParmList
    Exit Property
End Property
Property Get NamedParm() As Dictionary
    'NamedParmのリストを取得する
    Set NamedParm = dicNamedParm
    Exit Property
End Property
Public Function GetNamedList(ByVal strargParmName As String, ByVal longargDataType As BindType, ByVal varargOrigData As Variant) As Dictionary
    'NamedListのDictionaryを返す
    Dim varLocal As Variant
    Dim dicLocalNamedList As Dictionary
    Set dicLocalNamedList = New Dictionary
    Dim dicTypeAndData As Dictionary
    Set dicTypeAndData = New Dictionary
    On Error GoTo ErrorCatch
    If dicNamedParm Is Nothing Then
        Set dicNamedParm = New Dictionary
    End If
    If dicNamedParm.Exists(strargParmName) Then
        MsgBox "指定のバインドパラメーター名は既に存在するため、処理を中断します"
        Set GetNamedList = Nothing
        Exit Function
    End If
    'それぞれ指定されたデータ型にキャストしてやる
    Select Case longargDataType
    Case BindType.Int32
        varLocal = CLng(varargOrigData)
    Case BindType.Dbl
        varLocal = CDbl(varargOrigData)
    Case BindType.Text, BindType.Value
        varLocal = CStr(varargOrigData)
    End Select
    dicTypeAndData.Add Bind_Type, longargDataType
    dicTypeAndData.Add Bind_Data, varLocal
    dicLocalNamedList.Add strargParmName, dicTypeAndData
    Set GetNamedList = dicLocalNamedList
    Set dicLocalNamedList = Nothing
    Set dicTypeAndData = Nothing
    Exit Function
ErrorCatch:
    Debug.Print "GetNamedList code: " & Err.Number & " Description: " & Err.Description
    Set dicLocalNamedList = Nothing
    Set dicTypeAndData = Nothing
    Set GetNamedList = Nothing
    Exit Function
End Function
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
Public Function DoSQL_No_Transaction(Optional ByVal strSQLarg As String, Optional strargDBPath As String) As Boolean
    'SQL文字列とDBパス名（オプション）から、トランザクションなしの処理を行います（軽いの専用）
    Dim isCollect As Boolean
    Dim strErrorMsg As String
    On Error GoTo ErrorCatch
    'SQL指定されていればその内容で実行する
    '事前にプロパティで設定してもいいと思う
    If Not strSQLarg = "" Then
        Me.SQL = strSQLarg
    End If
    'DBパス指定されていれば、そちらに設定してやる
    If Not strargDBPath = "" Then
        Me.DBPath = strargDBPath
    End If
    'DBイニシャライズ・オープン
    isCollect = InitialAndOpenDB
    If Not isCollect Then
        '初期化かオープン失敗してるので終了
        DoSQL_No_Transaction = False
        Exit Function
    End If
    'SQL文字列コンパイル
    isCollect = Me.PrepareStmt
    If Not isCollect Then
        DoSQL_No_Transaction = False
'        GoTo NoTitleExit
    End If
    isCollect = Me.StepStmt
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
ErrorCatch:
    If Err.Number <> 0 Then
'        MsgBox Err.Number & vbCrLf & Err.Description
    End If
    Debug.Print "DoSQL_No_Transaction code: " & Err.Number & "description: " & Err.Description
    DoSQL_No_Transaction = False
    GoTo CloseAndExit
    Exit Function
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
NoTitleExit:
        myRecord = Array()
        ReDim myRecord(1)
        myRecord(0) = Array("No Title")
        myRecord(1) = Array(strErrorMsg)
        Exit Function
CloseAndExit:
    'StatamentHandle お掃除
    isCollect = Me.FinalizeStmt
    If Not isCollect Then
        MsgBox "DBステートメントファイナライズ失敗"
        DoSQL_No_Transaction = False
        Exit Function
    End If
    Exit Function
End Function
Private Function InitialAndOpenDB() As Boolean
    'DBの初期化とオープンを行う
    Dim isCollect As Boolean
    'まだDBイニシャライズしてないようならここでやる
    If Not isInitialize Then
        'DBイニシャライズ
        isCollect = DBInishialize()
        If Not isCollect Then
            '初期化失敗してるので終了
            Debug.Print "イニシャライズ失敗"
            MsgBox "DBイニシャライズ失敗"
            InitialAndOpenDB = False
            Exit Function
        End If
    End If
    'DB接続の有無を確認し、接続してなかったらここで接続する
    If Not isDBOpen Then
        isCollect = Me.OpenDB(Me.DBPath)
        If Not isCollect Then
            InitialAndOpenDB = False
            isDBOpen = False
            myRecord = Array()
            ReDim myRecord(1)
            myRecord(0) = Array("No Title")
            myRecord(1) = Array("open fail")
            Exit Function
        End If
    End If
    InitialAndOpenDB = True
    Exit Function
End Function
Public Function Do_SQL_With_Transaction(Optional ByVal strSQLarg As String, Optional strargDBPath As String) As Boolean
    'トランザクション処理ありのSQL実行
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
    isCollect = DoSQL_No_Transaction(strSQLarg, strargDBPath)
    If Not isCollect Then
        GoTo RollBackAndExit
    End If
    Do_SQL_With_Transaction = True
    '成功したのでコミットを
    Call CommitTransaction
    Exit Function
RollBackAndExit:
    'トランザクションをロールバックしたうえで、終了する
    Me.RollBackTransaction
    Do_SQL_With_Transaction = False
    Debug.Print "Do_SQL_With_Transaction code: " & Err.Number & "Description: " & Err.Description
    Exit Function
End Function
Public Function Do_SQL_Use_NamedParm_NO_Transaction() As Boolean
    'SQLとパラメータリストはプロパティで設定する
    Dim isCollect As Boolean
    If Me.SQL = "" Then
        MsgBox "SQL文字列が空なので実行を中断します"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    If Me.NamedParm Is Nothing Then
        MsgBox "バインドパラメーターが設定されていないため実行を中断します"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    'データベースイニシャライズ・オープン
    isCollect = InitialAndOpenDB
    If Not isCollect Then
        Debug.Print "Do_SQL_Use_NamedParm_NO_Transaction cant initialize DB"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    'SQlコンパイル
    isCollect = PrepareStmt()
    If Not isCollect Then
        Debug.Print "Do_SQL_Use_NamedParm_NO_Transaction Prepare fail"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    'パラメータバインド
    isCollect = BindingParm()
    If Not isCollect Then
        Debug.Print "Do_SQL_Use_NamedParm_NO_Transaction Binding fail"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    'Step（実際に実行される）
    isCollect = StepStmt()
    If Not isCollect Then
        Debug.Print "Step fail"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    'Reset_同じSQL分で、一部をループして条件を変え続けるような場合じゃないと不要・・？
    isCollect = ResetStmt(myStmtHandle)
    If Not isCollect Then
        Debug.Print "Reset fail"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    'Finalize(stmthandle お掃除)
    isCollect = FinalizeStmt()
    If Not isCollect Then
        Debug.Print "Finalize fail"
        Do_SQL_Use_NamedParm_NO_Transaction = False
        Exit Function
    End If
    Do_SQL_Use_NamedParm_NO_Transaction = True
    Exit Function
End Function
Public Function OpenDB(Optional ByVal strargDBFilePath As String) As Boolean
    Dim strErrorMsg As String
    If strargDBFilePath = "" And Me.DBPath = "" Then
        'どちらも指定されてなかったら・・・me.localmodeで処理分岐
        If Me.LocalMode Then
            Me.DBPath = ThisWorkbook.Path & "\" & LocalTempDBDir & "\" & LocalDBName
        Else
            Me.DBPath = constDatabasePath & "\" & constJobNumberDBname
        End If
    End If
    If Not strargDBFilePath = "" Then
        '引数で指定されてたら最優先にしてあげる
        strDBPath = strargDBFilePath
    End If
    'Job情報DBファイルオープン（無ければ作るはず）
    RetVal = SQLite3Open(Me.DBPath, myDbHandle)
    If RetVal <> SQLITE_OK Then
        '接続失敗
        strErrorMsg = "SQLite3Open returned" & RetVal & vbTab & SQLite3ErrMsg(myDbHandle)
        Debug.Print strErrorMsg
        Call ErrorOutput("OpenDB", strErrorMsg)
        isDBOpen = False
        OpenDB = False
        Exit Function
    End If
    '接続成功
    'DB接続フラグTrue
    isDBOpen = True
    OpenDB = True
    Exit Function
End Function
Public Function BeginTransaction() As Boolean
    Dim isCollect As Boolean
    Dim strErrorMsg
    isCollect = DoSQL_No_Transaction("BEGIN TRANSACTION")
    If Not isCollect Then
        strErrorMsg = "Begin Transaction Failed will RollBack"
        Debug.Print strErrorMsg
        Call ErrorOutput("BeginTransaction", strErrorMsg)
        BeginTransaction = False
        RollBackTransaction
        Exit Function
    End If
End Function
Public Function CommitTransaction() As Boolean
    'コミットトランザクション処理
    Dim isCollect As Boolean
    Dim strErrorMsg As String
    isCollect = DoSQL_No_Transaction("COMMIT TRANSACTION")
    If Not isCollect Then
        strErrorMsg = "Commit Failed will RollBack"
        Debug.Print strErrorMsg
        Call ErrorOutput("CommitTransaction", strErrorMsg)
        CommitTransaction = False
        RollBackTransaction
        Exit Function
    End If
End Function
Public Function RollBackTransaction() As Boolean
    'ロールバック処理
    DoSQL_No_Transaction ("ROLLBACK TRANSACTION")
    MsgBox "ロールバック発生"
    RollBackTransaction = True
    Exit Function
End Function
Public Function PrepareStmt() As Boolean
    Dim strErrorMsg
    'SQLステートメントコンパイル
    RetVal = SQLite3PrepareV2(Me.DbHandle, Me.SQL, myStmtHandle)
    If RetVal <> SQLITE_OK Then
        strErrorMsg = "SQLite3PrepareV2 returned" & RetVal & vbTab & SQLite3ErrMsg(myDbHandle)
        Debug.Print strErrorMsg
        Call ErrorOutput("PrepareStmt", strErrorMsg)
        PrepareStmt = False
        isStatement = False
        Exit Function
    End If
    isStatement = True
    PrepareStmt = True
    Exit Function
End Function
Public Function BindingParm() As Boolean
    'パラメータバインドを行う
    Dim varNamedKey As Variant
    Dim strErrorMsg As String
    Dim longParmIndex As Long
    For Each varNamedKey In Me.NamedParm
        longParmIndex = SQLite3BindParameterIndex(Me.StmtHandle, CStr(varNamedKey))
        If longParmIndex = 0 Then
            strErrorMsg = "BindingParm cant find ID parameter: " & CStr(varNamedKey)
            Debug.Print strErrorMsg
            Call ErrorOutput("BindingParm", strErrorMsg)
            BindingParm = False
            Exit Function
        End If
        Select Case Me.NamedParm(varNamedKey)(Bind_Type)
        'タイプによって挙動を変える
        Case BindType.Int32
            '整数値の場合(Long)
            RetVal = SQLite3BindInt32(Me.StmtHandle, longParmIndex, Me.NamedParm(varNamedKey)(Bind_Data))
        Case BindType.Dbl
            '小数値の場合(Double)
            RetVal = SQLite3BindDouble(Me.StmtHandle, longParmIndex, Me.NamedParm(varNamedKey)(Bind_Data))
        Case BindType.Text
            'テキストの場合(Text)
            RetVal = SQLite3BindText(Me.StmtHandle, longParmIndex, Me.NamedParm(varNamedKey)(Bind_Data))
        Case BindType.Blob
            'Blobの場合(Byte())
            RetVal = SQLite3BindBlob(Me.StmtHandle, longParmIndex, Array((Me.NamedParm(varNamedKey)(Bind_Data))))
        Case BindType.Nul
            'Null1の場合
            RetVal = SQLite3BindNull(Me.StmtHandle, longParmIndex)
        Case BindType.Value
            'Valueの場合
            RetVal = SQLite3BindValue(Me.StmtHandle, longParmIndex, Me.NamedParm(varNamedKey)(Bind_Data))
        Case Else
            'それ以外・・は普通は入らないはずなんだけど・・・
            MsgBox "指定外のタイプが入力されました。どうして？"
            BindingParm = False
            Call ErrorOutput("BindingParm", "指定外のタイプが入力されました")
            Exit Function
        End Select
        If RetVal <> SQLITE_OK Then
            strErrorMsg = "Sqlite3Bind returned " & RetVal & SQLite3ErrMsg(Me.DbHandle)
            Debug.Print strErrorMsg
            Call ErrorOutput("BindingParm", strErrorMsg)
            BindingParm = False
            Exit Function
        End If
    Next varNamedKey
    BindingParm = True
    Exit Function
End Function
Public Function StepStmt() As Boolean
    'SQL実行（パラメータバインドの場合はここで本実行）
    Dim strErrorMsg As String
    Dim isCollect As Boolean
    RetVal = SQLite3Step(Me.StmtHandle)
    If RetVal = SQLITE_OK Or RetVal = SQLITE_DONE Then
        'ここで正常終了
        'Insertとか結果が一つのはこっちになるはず
        'Sqlite3Changesで結果行数取得してそれで終わり
        strErrorMsg = "直前の操作で " & SQLite3Changes(myDbHandle) & " 個の変更がありました（正常終了）"
        Debug.Print strErrorMsg
        StepStmt = True
        GoTo NoTitleExit
        Exit Function
    ElseIf RetVal = SQLITE_ROW Then
        'SQL実行とデータ配列格納をここからやる
        isCollect = Get_Data_and_Put_Arry
        If Not isCollect Then
            Debug.Print "StepStmt 配列格納時に何かあったらしい"
            Call ErrorOutput("StepStmt_Put_Arry", "Some happen....")
            StepStmt = False
            Exit Function
        End If
    Else
        'ここはもう異常終了でいいと思う
        strErrorMsg = "SQLite3Step returned " & RetVal & SQLite3ErrMsg(myDbHandle)
        Debug.Print strErrorMsg
        Call ErrorOutput("StepStmt", strErrorMsg)
        StepStmt = False
        Exit Function
    End If
    '処理成功した場合はここ
    StepStmt = True
    Exit Function
NoTitleExit:
        myRecord = Array()
        ReDim myRecord(1)
        myRecord(0) = Array("StepStmt")
        myRecord(1) = Array(strErrorMsg)
        Exit Function
End Function
Public Function ResetStmt(Optional ByVal longptrargStmtHandle As LongPtr) As Boolean
    Dim strErrorMsg As String
    'バインドパラメータ使用時、次回パラメータセットのためのリセット処理
    RetVal = SQLite3Reset(Me.StmtHandle)
    If RetVal <> SQLITE_OK Then
        strErrorMsg = "SQLite3Reset returned " & RetVal & SQLite3ErrMsg(myDbHandle)
        Debug.Print strErrorMsg
        Call ErrorOutput("ResetStmt", strErrorMsg)
        ResetStmt = False
        Exit Function
    End If
    ResetStmt = True
    Exit Function
End Function
Public Function FinalizeStmt() As Boolean
    Dim strErrorMsg As String
    'Stmtお掃除
    If Not Me.StmtHandle = 0 Then
        RetVal = SQLite3Finalize(Me.StmtHandle)
        If RetVal <> SQLITE_OK Then
            strErrorMsg = "SQLite3Finalize returned" & RetVal & vbTab & SQLite3ErrMsg(myDbHandle)
            'Debug.Print "SQLite3Finalize returned" & RetVal, SQLite3ErrMsg(myDbHandle)
            Debug.Print strErrorMsg
            Call ErrorOutput("FinalizeStmt", strErrorMsg)
            FinalizeStmt = False
            Exit Function
        End If
        FinalizeStmt = True
        Me.StmtHandle = 0
        isStatement = False
        Exit Function
    End If
    FinalizeStmt = True
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
    Dim strErrorMsg As String
    Dim dblStartTime As Double
    'SQL実行
    On Error GoTo ErrorCatch
    If (RetVal <> SQLITE_OK) And (RetVal <> SQLITE_DONE) And (RetVal <> SQLITE_ROW) Then
        '異常値なので即終了
        strErrorMsg = "SQLite3Step returned" & RetVal & SQLite3ErrMsg(myDbHandle)
        Debug.Print strErrorMsg
        Call ErrorOutput("Put_Arry", strErrorMsg)
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
Private Function ColumnValue_InClass(ByVal StmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#Else
Private Function ColumnValue_InClass(ByVal StmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#End If
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue_InClass = SQLite3ColumnInt32(StmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue_InClass = SQLite3ColumnDouble(StmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue_InClass = SQLite3ColumnText(StmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue_InClass = SQLite3ColumnText(StmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue_InClass = Null
    End Select
End Function
Private Function DBInishialize(Optional ByRef strLibDir As String) As Boolean
    #If Win64 Then
        If strLibDir = "" Then
            'Libディレクトリ指定されなかった場合は、DBディレクトリと同一とする（お引越し楽だから・・・）→ネットワーク越しがあまりにも遅いためx64サブディレクトリへ
            strLibDir = ThisWorkbook.Path + "\x64"
        End If
    #Else
        If strLibDir = "" Then
            '32ビットの場合はブックのカレントディレクトリ
            strLibDir = ThisWorkbook.Path
        End If
    #End If
    'イニシャライズ
    Dim InitReturn As Long
    Dim isCollect
    #If Win64 Then
        ' I put the 64-bit version of SQLite.dll under a subdirectory called x64
        InitReturn = SQLite3Initialize(strLibDir)
    #Else
        InitReturn = SQLite3Initialize ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    #End If
    If InitReturn = SQLITE_INIT_ERROR Then
        MsgBox "システムファイルの初期化に失敗しました。ネットワーク接続が原因の場合もあります。処理を中断します。"
        DBInishialize = False
        Exit Function
    End If
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        DBInishialize = False
        Exit Function
    End If
    DBInishialize = True
    'DBディレクトリに移動
'    isCollect = ChcurrentforDB()
'    isCollect = ChCurrentToDBDirectory
'    If Not isCollect Then
'        MsgBox "DBディレクトリ存在確認中にエラー発生"
'        DBInishialize = False
'        Exit Function
'    End If
    DBInishialize = True
    isInitialize = True
End Function
Private Function ErrorOutput(ByVal strargTitle As String, ByVal strargErrorMsg As String) As Boolean
    'エラーが出た時に、結果レコードにその旨設定してやる
    Erase myRecord
    myRecord = Array()
    ReDim myRecord(1)
    myRecord(0) = Array(strargTitle)
    myRecord(1) = Array(strargErrorMsg)
    Exit Function
End Function