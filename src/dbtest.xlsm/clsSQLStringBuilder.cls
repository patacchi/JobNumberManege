VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'SQL文字操作（主に結合）のためのクラス
'開始履歴、終了履歴を設定して、文字列連結は配列のJoinで行う方向
Private strSQLBuilder As String         '最後の結果のSQL文
Private arrstrField As Variant          'フィールドの文字列配列
Private intFieldCount As Integer        'フィールドの数
Private longBulkCount As Long            'バルク数
Private strStartRireki As String        'スタートの履歴
Private longMaisuu As Long              '枚数
Private byteRenbanketa As Byte          '連番（数字部分）の桁数
Private strEndRireki As String          'エンドの履歴
Private longStringLength As Long        '文字列の長さ
Private strTableName As String          'テーブル名
Private strJobNumber As String          'ジョブ番号
Private dicFieldType As Dictionary      'フィールドのタイプ名
Private strRirekiHeader As String       '履歴ヘッダ
Private longRirekiNumber As Long        '履歴連番部分
Private dblTimer As Double              '処理時間計測用タイマーアーアーア
'Private longRecordCount As Long         '現在の処理中データの場所（バルクの時に使う）
Private Sub Class_Initialize()
    'クラス初期化動作
    'なぜかここでフィールドリストの設定を行う
    arrFieldList_JobData = Split(Job_Number & "," & Job_RirekiHeader & "," & Job_RirekiNumber & "," & Job_Rireki & "," & Field_Initialdate, ",")
    arrFieldList_Barcode = Split(Field_BarcordeNumber & "," & Laser_Rireki & "," & Field_Initialdate, ",")
    arrFieldList_Retry = Split(Field_BarcordeNumber & "," & Retry_Rireki & "," & Retry_Reason & "," & Field_Initialdate, ",")
    strSQLBuilder = Empty
    intFieldCount = 0
'    longRecordCount = 0
    longBulkCount = 1000000                  'バルク数初期設定
    strStartRireki = Empty
    strEndRireki = Empty
    longStringLength = 0
End Sub
Private Sub Class_Terminate()
    'クラス破棄時に実行
    strSQLBuilder = Empty
    intFieldCount = 0
    strStartRireki = Empty
    strEndRireki = Empty
    longStringLength = 0
End Sub
'プロパティ設定
Property Let FieldArray(varargField As Variant)
    'フィールド名を配列の形で指定
    arrstrField = varargField
End Property
Property Get FieldArray() As Variant
    'フィールドの配列を返す
    FieldArray = arrstrField
End Property
Property Let BulkCount(longargBulkCount As Long)
    'バルク数設定
    longBulkCount = longargBulkCount
End Property
Property Get BulkCount() As Long
    'バルク数取得
    BulkCount = longBulkCount
End Property
Property Let StartRireki(strargStartRireki As String)
    'スタート履歴設定
    strStartRireki = strargStartRireki
End Property
Property Get StartRireki() As String
    'スタート履歴取得
    StartRireki = strStartRireki
End Property
Property Let EndRireki(strargEndRireki As String)
    'エンド履歴設定
    strEndRireki = strargEndRireki
End Property
Property Get EndRireki() As String
    'エンド履歴取得
    EndRireki = strEndRireki
End Property
Property Let TableName(strargTableName As String)
    'テーブル名設定
    strTableName = strargTableName
End Property
Property Get TableName() As String
    'テーブル名取得
    TableName = strTableName
End Property
Property Let Maisu(longargMaisuu As Long)
    '枚数設定
    longMaisuu = longargMaisuu
End Property
Property Get Maisu() As Long
    '枚数取得
    Maisu = longMaisuu
End Property
Property Let RenbanKeta(byteargKetasuu As Byte)
    '連番桁数設定
    byteRenbanketa = byteargKetasuu
End Property
Property Get RenbanKeta() As Byte
    '連番桁数取得
    RenbanKeta = byteRenbanketa
End Property
Property Let JobNumber(strargJobNumber As String)
    'ジョブ番号設定
    strJobNumber = strargJobNumber
End Property
Property Get JobNumber() As String
    'ジョブ番号取得
    JobNumber = strJobNumber
End Property
Property Set FieldType(ByRef dicargFieldType As Dictionary)
    'フィールドタイプ設定
    Set dicFieldType = dicargFieldType
End Property
Property Get FieldType() As Dictionary
    'フィールドタイプ取得
    Set FieldType = dicFieldType
End Property
Property Let RirekiHeader(strargRirekiHeader As String)
    '履歴ヘッダ設定
    strRirekiHeader = strargRirekiHeader
End Property
Property Get RirekiHeader() As String
    '履歴ヘッダ取得
    RirekiHeader = strRirekiHeader
End Property
Property Let RirekiNumber(longargRirekiNumber As Long)
    '履歴の連番設定
    longRirekiNumber = longargRirekiNumber
End Property
Property Get RirekiNumber() As Long
    '履歴の連番取得
    RirekiNumber = longRirekiNumber
End Property
Property Get Rireki() As String
'    '履歴（結合したの）
    Rireki = Me.RirekiHeader & Right(String$(Me.RenbanKeta, "0") & Me.RirekiNumber, Me.RenbanKeta)
End Property
'メソｄｄッド
Public Function CreateInsertSQL(Optional ByVal boolCheckLastRireki As Boolean, _
                                Optional ByVal strargStartRireki As String, _
                                Optional ByVal longargMaisuu As Long, _
                                Optional ByVal strargTableName As String, _
                                Optional ByVal strargJobNumber As String) As Boolean
    Dim KishuInfo As typKishuInfo
    Dim longRirekiCount As Long
    Dim strField As String          'フィールド部分までのSQL
    Dim intReturnCode               'msgbox関数のリターンコード
    On Error GoTo ErrorCatch
    'フィールド名は先にプロパティでセットしてほしいの・・・
    If Not strargStartRireki = "" Then
        Me.StartRireki = strargStartRireki
    End If
    If Not longargMaisuu = 0 Then
        Me.Maisu = longargMaisuu
    End If
    If Not strargTableName = "" Then
        Me.TableName = strargTableName
    End If
    If Not strargJobNumber = "" Then
        Me.JobNumber = strargJobNumber
    End If
    If Me.StartRireki = "" Or _
        Me.Maisu <= 0 Or _
        Me.TableName = "" Or _
        Me.JobNumber = "" Then
        MsgBox "CreateInsertSQL エラー：必須項目で入力されていない場所があります"
        CreateInsertSQL = False
        Exit Function
    End If
    'DBから履歴の桁数とか拾ってくる
    KishuInfo = getKishuInfoByRireki(Me.StartRireki)
    If Me.RenbanKeta = 0 Then
        Me.RenbanKeta = KishuInfo.RenbanKetasuu
    End If
    Me.RirekiHeader = Left(Me.StartRireki, KishuInfo.TotalRirekiketa - KishuInfo.RenbanKetasuu)
    'FieldType定義がなされているかチェック
    If dicFieldType Is Nothing Then
        Set dicFieldType = GetFieldTypeNameByTableName(Me.TableName)
    End If
    'ラスト履歴チェックの有無により処理分岐
    If boolCheckLastRireki Then
        Me.EndRireki = Me.RirekiHeader & Right(String$(Me.RenbanKeta, "0") & _
                        CLng(Right(Me.StartRireki, KishuInfo.RenbanKetasuu)) + Me.Maisu - 1, Me.RenbanKeta)
        intReturnCode = MsgBox(prompt:="最終履歴は " & Me.EndRireki & " で合っていますか？", Buttons:=vbYesNo)
        If intReturnCode = vbNo Then
            Debug.Print "Last Rireki Check Abort"
            CreateInsertSQL = False
            Exit Function
        End If
    End If
    dblTimer = timer()
    Debug.Print "CreateInsertSQL バルク対応インサートSQL作成・処理開始"
    Call BuildFieldList
    Call BuildJobDataValuesList
    CreateInsertSQL = True
    Exit Function
ErrorCatch:
    Debug.Print "CreateInsertSQL: " & Err.Number & "Description: " & Err.Description
End Function
Private Function BuildFieldList() As Boolean
    Dim intValueCount As Integer
    Dim strarrBuf() As String
    Dim intCounter As Integer
    On Error GoTo ErrorCatch
    If Not strSQLBuilder = "" Then
        Debug.Print "Already string here"
        Exit Function
    End If
    '最初の1発目は項目入れちゃうので、で、最後もやっぱりカッコ入れるので、数が変わらない・・・
    intValueCount = UBound(Me.FieldArray) - LBound(Me.FieldArray)
    ReDim strarrBuf(intValueCount)
    For intCounter = 0 To intValueCount
        Select Case intCounter
        Case 0
            '最初
            strarrBuf(intCounter) = "INSERT INTO " & Me.TableName & " (" & Me.FieldArray(intCounter)
        Case intValueCount
            '最後
            strarrBuf(intCounter) = Me.FieldArray(intCounter) & ") VALUES "
        Case Else
            '途中
            strarrBuf(intCounter) = Me.addQuote(Me.FieldArray(intCounter))
        End Select
    Next intCounter
    '文字結合実行
    strSQLBuilder = Append(strSQLBuilder, strarrBuf)
    BuildFieldList = True
    Exit Function
ErrorCatch:
    Debug.Print "BuildFiedList code: " & Err.Number & "Description: " & Err.Description
End Function
Private Function BuildJobDataValuesList() As Boolean
    'INSERT のVALUES 以降の続きを構築する
    Dim longCounterBulk As Long           'こいつがme.BulkCountで割り切れたら、いったんそこでSQL発行する
    Dim intFieldCounter As Integer
    Dim longMaisuuCounter As Long
    Dim longRirekiNumber As Long
    Dim strarrBuf() As String
    Dim strValues(0) As String
    Dim longCurrentElement As Long          'バッファの現在の（通算）要素数、フィールド×レコード数になるかも
    Dim isCollect As Boolean
    Dim dbJobData As clsSQLiteHandle
    Dim strLocalTimeWithMSec As String      'ミリ秒付きの現在日時
    On Error GoTo ErrorCatch
'    Set dbJobData = New clsSQLiteHandle
    longCurrentElement = 0
    strValues(0) = ""
    '現在日時を取得
    strLocalTimeWithMSec = GetLocalTimeWithMilliSec()
    'なぜかもう一度実行されてしまう事があるので、先頭で続行可否を判断
'    If Me.Maisu <= longRecordCount Then
'        'もう必要枚数は終わってるので、処理を中断して抜ける
'        BuildJobDataValuesList = False
'        GoTo CloseAndExit
'        Exit Function
'    End If
    'バルク数か枚数の少ないほう × フィールド数分のバッファを確保
    ReDim strarrBuf(Application.WorksheetFunction.Min _
    (Me.BulkCount, Me.Maisu) * (UBound(Me.FieldArray) - LBound(Me.FieldArray) + 1) - 1)
    longMaisuuCounter = 1
    longCounterBulk = 1
    'これが一番外側のループ。Forはだめ。指定された枚数が終わるまで繰り返す
    Do While longMaisuuCounter <= Me.Maisu
        'このループを抜けたらSQL実行し、SQLクリアしてまたやりなおす
        'バルクカウントだけど、これもForはだめ、中で数が変わるので
        Do While longCounterBulk <= Me.BulkCount
            'ここで履歴の連番を導出する
            Me.RirekiNumber = CLng(Right(Me.StartRireki, Me.RenbanKeta)) + longMaisuuCounter - 1
            '一番内側のループ。ここでフィールドの数分ループする（文法のひとまとまり）
            For intFieldCount = LBound(Me.FieldArray) To UBound(Me.FieldArray)
                'それぞれフィールド名によって導出方法が異なる
                Select Case Me.FieldArray(intFieldCount)
                Case Job_Number
                    strarrBuf(longCurrentElement) = Me.JobNumber
                Case Job_RirekiHeader
                    strarrBuf(longCurrentElement) = Me.RirekiHeader
                Case Job_RirekiNumber
                    strarrBuf(longCurrentElement) = Me.RirekiNumber
                Case Job_Rireki
                    strarrBuf(longCurrentElement) = Me.Rireki
                Case Field_Initialdate
                    strarrBuf(longCurrentElement) = strLocalTimeWithMSec
                Case Else
                    MsgBox Me.FieldArray(intFieldCount) & "は知らない子です"
                End Select
                'データ型によって、クオートの必要性の有無を判断する
                Select Case Me.FieldType.Item(Me.FieldArray(intFieldCount))
                Case "TEXT", "BLOB"
                    'クオートが必要
                    strarrBuf(longCurrentElement) = addQuote(strarrBuf(longCurrentElement))
                Case "NUMERIC", "INTEGER", "FLOAT"
                    'とりあえずそのままで
                End Select
                'フィールドループの最後なので、カッコだのの処理をしてあげる
                If intFieldCount = LBound(Me.FieldArray) Then
                    'Valueの最初なので(を先頭につけてあげる
                    strarrBuf(longCurrentElement) = "(" & strarrBuf(longCurrentElement)
                ElseIf intFieldCount = UBound(Me.FieldArray) Then
                    '最後なので、カッコ閉じてあげる )
                    strarrBuf(longCurrentElement) = strarrBuf(longCurrentElement) & ")"
                End If
                longCurrentElement = longCurrentElement + 1
            Next intFieldCount
            '終わったら枚数とバルクカウンターインクリメント
            longMaisuuCounter = longMaisuuCounter + 1
            longCounterBulk = longCounterBulk + 1
            'バルクカウントとは別に、毎回枚数も監視すること、上限に達したらすぐに抜ける
            If longCounterBulk - 1 Mod Me.BulkCount = 0 Or _
                longMaisuuCounter > Me.Maisu Then
                Exit Do
            End If
        Loop
        '規定数に達したので、文字結合の上、SQL実行し、SQL文クリア、フィールド再構築をする
        '今のところとりあえず方法がないので、Unique違反があった行は空いてるUpdateDateのフィールドに現在時刻を突っ込む
        'ループ抜けたら文字連結
        strarrBuf(longCurrentElement - 1) = strarrBuf(longCurrentElement - 1) & vbCrLf & " on CONFLICT(RirekiNumber) do UPDATE set UpdateDate = CURRENT_TIMESTAMP"
        Debug.Print "配列格納完了。バルク数 " & Me.BulkCount & " 、要素数 " & longCurrentElement & " 個で " & timer - dblTimer & " 秒掛かりました。"
        Set dbJobData = New clsSQLiteHandle
        'まずはValuesの中を連結してやる
        strValues(0) = Append("", strarrBuf)
        '次に、Fieldsの部分と連結
        strSQLBuilder = Append(strSQLBuilder, strValues, " ")
        dbJobData.SQL = strSQLBuilder
        Debug.Print "文字連結完了。バルク数 " & Me.BulkCount & " で " & timer - dblTimer & " 秒掛かりました。"
        'SQL実行
        isCollect = dbJobData.Do_SQL_With_Transaction()
        If Not isCollect Then
            Debug.Print "バルクインサート中に何かあったようですにょ"
        End If
        strSQLBuilder = ""
        strValues(0) = ""
        longCurrentElement = 0
        longCounterBulk = 1
        '毎回接続を破棄してみる
        Set dbJobData = Nothing
        If Me.Maisu <= longMaisuuCounter Then
            'もう必要枚数は終わってるので、処理を中断して抜ける
            BuildJobDataValuesList = True
            GoTo CloseAndExit
            Exit Function
        End If
        'バルク数か枚数の残りの少ないほう × フィールド数分のバッファを確保
        ReDim strarrBuf(Application.WorksheetFunction.Min _
        (Me.BulkCount, (Me.Maisu - longMaisuuCounter + 1)) * (UBound(Me.FieldArray) - LBound(Me.FieldArray) + 1) - 1)
        '次の周回のSQL構築準備をする
        Call BuildFieldList
    Loop
    BuildJobDataValuesList = True
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
    Debug.Print "多分処理完了 " & "バルク数:" & Me.BulkCount & "、枚数:" & Me.Maisu & "枚で " & timer() - dblTimer & " 秒掛かったようです"
    Set dbJobData = Nothing
    Exit Function
ErrorCatch:
    Debug.Print "BuildJobDataValues code: " & Err.Number & "Description: " & Err.Description
End Function
Private Function Append(ByVal strargOrignString As String, ByRef varAryString As Variant, Optional strDelimiter As String = ",") As String
    Dim strarrBuf() As String
    Dim longElementCount As Long
    Dim longCounter As Long
    On Error GoTo ErrorCatch
    Select Case Len(strargOrignString)
    Case 0
        '元のStringが長さ0の時は、追加の分が最初
        longElementCount = UBound(varAryString) - LBound(varAryString)
    Case Else
        '元のが文字ありだった場合
        '引数に指定されている要素数 + オリジナルのプラス一つ
        longElementCount = UBound(varAryString) - LBound(varAryString) + 1
    End Select
    ReDim strarrBuf(longElementCount)
    Select Case Len(strargOrignString)
    Case 0
        '元の長さが0の場合は追加の配列と同じでいい・・・
        strarrBuf = varAryString
    Case Else
        'オリジナルが文字ありだった場合
        'まず先頭にオリジナルのStringをくっつけてやる
        longCounter = 0
        strarrBuf(longCounter) = strargOrignString
        '追加分の配列を順番にくっつける
        For longCounter = 1 To longElementCount
            strarrBuf(longCounter) = varAryString(longCounter - (1 - LBound(varAryString)))
        Next longCounter
    End Select
    '配列に格納し終わったらJoinでStringにする
    Append = Join(strarrBuf, strDelimiter)
    Exit Function
ErrorCatch:
    Debug.Print "Append Error code: " & Err.Number & "Description: " & Err.Description
End Function
Public Function addQuote(ByVal strargOriginalString As String, Optional ByVal strQuote As String = """") As String
    '与えられた文字列に、エスケープ処理をする
    Dim strBuf(2) As String
    On Error GoTo ErrorCatch
    strBuf(0) = strQuote
    strBuf(1) = Replace(strargOriginalString, """", """""")
    strBuf(2) = strQuote
    addQuote = Append("", strBuf, "")
    Exit Function
ErrorCatch:
    Debug.Print "addQuote Error code: " & Err.Number & "Description: " & Err.Description
End Function