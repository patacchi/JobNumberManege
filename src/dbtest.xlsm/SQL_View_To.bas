Attribute VB_Name = "SQL_View_To"
Option Explicit
'将来的にViewテーブルにしたいSQLを集めます
Public Function ReturnJobNumber_For_KanbanDivide(ByVal strargTableName As String, ByRef argKishuInfo As typKishuInfo) As Variant
    Dim strSQL As String
    Dim strRirekiNumber As String
    Dim strFromTable As String
    Dim strJobNumber As String
    Dim strInitialDate As String
    Dim strKanbanChr As String
    Dim longStartNumber_Temp As Long                '履歴連番の暫定スタート
    Dim jobStart As typJobInfo                      'スタート連番のJobInfo
    Dim longStartRirekiNumber As Long               '履歴連番のスタート
    Dim longEndNumber_Temp As Long                  '履歴連番の暫定エンド
    Dim jobEnd As typJobInfo                        'エンド連番のJobInfo
    Dim longEndRirekiNumber As Long                 '履歴連番のエンド
    Dim longKanbanCurrent As Long                   '看板分割の現在の最終連番
    Dim longJobCurrent As Long                      'Job登録の最終連番
    Dim vararrKanbanJobNumber As Variant
    Dim dbKanbanJobNumber As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Set dbKanbanJobNumber = New clsSQLiteHandle
    #If READLOCAL Then
        dbKanbanJobNumber.LocalMode = True
    #Else
        dbKanbanJobNumber.LocalMode = False
    #End If
    Set sqlbC = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    '以下をパラメータバインド、かつWith句で取得件数制限する
    'パラメータバインドにする変数はaddquoteしちゃダメ→GroupByにパラメータ効かなかった・・・
    strRirekiNumber = sqlbC.addQuote(Job_RirekiNumber)
'    strRirekiNumber = Job_RirekiNumber
    strFromTable = sqlbC.addQuote(strargTableName)
'    strFromTable = strargTableName
    strJobNumber = sqlbC.addQuote(Job_Number)
'    strJobNumber = Job_Number
    strInitialDate = sqlbC.addQuote(Field_Initialdate)
'    strInitialDate = Field_Initialdate
    strKanbanChr = sqlbC.addQuote(Job_KanbanChr)
'    strKanbanChr = Job_KanbanChr
    '看板とJob登録、それぞれの連番のエンドを取得
    longKanbanCurrent = GetLastRirekiNumber_byKishuTable(argKishuInfo.KishuName, KanbanChrField, boolRenumberIfZero:=True)
    longJobCurrent = GetLastRirekiNumber_byKishuTable(argKishuInfo.KishuName, JobNumberField, True)
    '連番のスタート・エンドを求める
    Select Case frmKanban.chkBoxLastArea
    Case True
        '範囲制限が設定されている場合（こっちがデフォルト）
        '暫定のスタート・エンドを求める
        longStartNumber_Temp = Application.WorksheetFunction.Max(1, longKanbanCurrent - CLng(frmKanban.txtBoxBeforeArea.Text))
        longEndNumber_Temp = Application.WorksheetFunction.Min(longJobCurrent, longKanbanCurrent + CLng(frmKanban.txtBoxAfterArea.Text))
    Case False
        '範囲制限なしの場合
        longStartNumber_Temp = 1
        longEndNumber_Temp = longJobCurrent
    End Select
    '求められた連番から、Job単位にRoudUp,RoundDownを行う
    jobStart = GetRoundJobInfo_byRirekiNumber(longStartNumber_Temp, Floor_to_Ceil, argKishuInfo)
    jobEnd = GetRoundJobInfo_byRirekiNumber(longEndNumber_Temp, Ceil_to_Floor, argKishuInfo)
    longStartRirekiNumber = jobStart.StartNumber
    longEndRirekiNumber = jobEnd.EndNumber
    'バインド用リスト設定
    Set dbKanbanJobNumber.NamedParm = dbKanbanJobNumber.GetNamedList("@RirekiStart", Int32, longStartRirekiNumber)
    Set dbKanbanJobNumber.NamedParm = dbKanbanJobNumber.GetNamedList("@RirekiEnd", Int32, longEndRirekiNumber)
'    Set dbKanbanJobNumber.NamedParm = dbKanbanJobNumber.GetNamedList("@GroupBy1", Text, strJobNumber)
'    Set dbKanbanJobNumber.NamedParm = dbKanbanJobNumber.GetNamedList("@GroupBy2", Text, strInitialDate)
'    Set dbKanbanJobNumber.NamedParm = dbKanbanJobNumber.GetNamedList("@GroupBy3", Text, strKanbanChr)
'    With_Local
'    With_Remote
'    strSQL = strSQL & "SELECT " & sqlbC.addQuote(Job_Number) & " as ""Job番号"", " & sqlbC.addQuote(Field_Initialdate) & " as ""登録日時"",count(*) - count("
'    strSQL = strSQL & sqlbC.addQuote(Job_KanbanChr) & ") as ""残り枚数"" FROM " & sqlbC.addQuote(strargTableName)
''    strSQL = strSQL & " WHERE " & sqlbC.addQuote(Job_KanbanChr) & " IS NULL GROUP BY " & sqlbC.addQuote(Job_Number) & "," & sqlbC.addQuote(Field_Initialdate)
'    strSQL = strSQL & " GROUP BY " & sqlbC.addQuote(Job_Number) & "," & sqlbC.addQuote(Field_Initialdate)
'    strSQL = strSQL & " ORDER BY " & sqlbC.addQuote(Job_RirekiNumber) & " ASC;"
    strSQL = "WITH " & With_Remote & " AS (SELECT * FROM " & strFromTable & " WHERE " & strRirekiNumber & " BETWEEN @RirekiStart AND @RirekiEnd)"
    strSQL = strSQL & "SELECT " & strJobNumber & " as ""Job番号""," & strInitialDate & " as ""登録日時"",count(*) - count("
    strSQL = strSQL & strKanbanChr & ") as ""残り枚数"" FROM " & With_Remote
    'strSql = strSql & " GROUP BY @GroupBy1,@GroupBy2,@GroupBy3 ORDER BY " & sqlbC.addQuote(strRirekiNumber) & " ASC;"
    strSQL = strSQL & " GROUP BY " & strJobNumber & "," & strInitialDate & " ORDER BY " & strRirekiNumber & " ASC;"
    dbKanbanJobNumber.SQL = strSQL
'    dbKanbanJobNumber.DoSQL_No_Transaction
    dbKanbanJobNumber.Do_SQL_Use_NamedParm_NO_Transaction
    vararrKanbanJobNumber = dbKanbanJobNumber.RS_Array(boolPlusTytle:=False)
    Set dbKanbanJobNumber = Nothing
    Set sqlbC = Nothing
    ReturnJobNumber_For_KanbanDivide = vararrKanbanJobNumber
    Exit Function
ErrorCatch:
    Set dbKanbanJobNumber = Nothing
    Set sqlbC = Nothing
    Debug.Print "ReturnJobNumber code: " & Err.Number & " Description: " & Err.Description
    Exit Function
End Function
Public Function ReturnJobInfo_by_JobNumber(ByRef argKishuInfo As typKishuInfo, ByVal strargJobNumber As String, ByVal strargInitialDate As String) As typJobInfo
    'Job番号とInitialInputDateを与えてもらって、TypJobInfo型で返す
    '今のところはスタートとエンドの連番・履歴のみ
    Dim localJobInfo As typJobInfo
    Dim strSQL As String
    Dim strTableName As String
    Dim strFieldJobNumber As String
    Dim strFieldRireki As String
    Dim strFieldInitialDate As String
    Dim varReturn As Variant
    Dim dbJobInfo As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    On Error GoTo ErrorCatch
    Set sqlbC = New clsSQLStringBuilder
    Set dbJobInfo = New clsSQLiteHandle
    strTableName = sqlbC.addQuote(Table_JobDataPri & argKishuInfo.KishuName)
    strFieldJobNumber = sqlbC.addQuote(Job_Number)
    strFieldRireki = sqlbC.addQuote(Job_Rireki)
    strFieldInitialDate = sqlbC.addQuote(Field_Initialdate)
    Set dbJobInfo.NamedParm = dbJobInfo.GetNamedList("@JobNumber", Text, strargJobNumber)
    Set dbJobInfo.NamedParm = dbJobInfo.GetNamedList("@InitialDate", Text, strargInitialDate)
'    WITH Remote_Limit as (select * FROM remote.T_JobData_Test15 WHERE JobNumber = "TT00122" AND InitialInputDate = "2021-03-19T23:09:33.175")
'    SELECT MIN(rireki),MAX(rireki)  FROM Remote_Limit;
    strSQL = "WITH " & With_Remote & " AS (SELECT * FROM " & strTableName & " WHERE " & strFieldJobNumber & " = @JobNumber AND "
    strSQL = strSQL & strFieldInitialDate & " = @InitialDate) "
    strSQL = strSQL & " SELECT MIN(" & strFieldRireki & "),MAX(" & strFieldRireki & ")  FROM " & With_Remote & ";"
    dbJobInfo.SQL = strSQL
    Call dbJobInfo.Do_SQL_Use_NamedParm_NO_Transaction
    varReturn = dbJobInfo.RS_Array
    Set dbJobInfo = Nothing
    Set sqlbC = Nothing
    If UBound(varReturn, 2) < 1 Then
        '結果取得に失敗してるぽい
        ReturnJobInfo_by_JobNumber = localJobInfo
        Exit Function
    End If
    localJobInfo.startRireki = varReturn(0, 0)
    localJobInfo.EndRireki = varReturn(0, 1)
    localJobInfo.StartNumber = CLng(Right(localJobInfo.startRireki, argKishuInfo.RenbanKetasuu))
    localJobInfo.EndNumber = CLng(Right(localJobInfo.EndRireki, argKishuInfo.RenbanKetasuu))
    localJobInfo.JobNumber = strargJobNumber
    localJobInfo.InitialDate = strargInitialDate
    ReturnJobInfo_by_JobNumber = localJobInfo
    Exit Function
ErrorCatch:
    Debug.Print "ReturnJobInfo_by_JobNumber code: " & Err.Number & " Description: " & Err.Description
    Exit Function
End Function
Public Function GetRoundJobInfo_byRirekiNumber(ByVal longargRirekiNumber As Long, ByVal longFindDir As RirekiFindDir, ByRef argKishuInfo As typKishuInfo) As typJobInfo
    Dim strTableName As String
    Dim longCurrentNumber As Long
    Dim longMinNumber As Long
    Dim longMaxNumber As Long
    Dim longFindNumber As Long      'お探しの番号はこれですか
    Dim JobInfoLocal As typJobInfo
    On Error GoTo ErrorCatch
    strTableName = Table_JobDataPri & argKishuInfo.KishuName
    'それぞれの連番を求めていく
    longCurrentNumber = longargRirekiNumber
    longMinNumber = GetMinimumRirekiNumber_byTableName(strTableName)
    longMaxNumber = GetLastRirekiNumber_byKishuTable(argKishuInfo.KishuName, JobNumberField, True)
    '番号を決定する
    Select Case longFindDir
    Case RirekiFindDir.Ceil_to_Floor
        If longCurrentNumber > longMaxNumber Then
            'maxより大きいのはだめぇ・・・
            longFindNumber = longMaxNumber
        Else
            'max以下ならそのまま使う
            longFindNumber = longCurrentNumber
        End If
    Case RirekiFindDir.Floor_to_Ceil
        If longCurrentNumber < longMinNumber Then
            'min未満はだめ
            longFindNumber = longMinNumber
        Else
            'min以上ならそのまま使う
            longFindNumber = longCurrentNumber
        End If
    End Select
    JobInfoLocal = GetJobInfo_By_RirekiNumberandKishuInfo(longFindNumber, argKishuInfo)
    GetRoundJobInfo_byRirekiNumber = JobInfoLocal
    Exit Function
ErrorCatch:
    Debug.Print "GetRoundJobInfo_ByRirekiNumber code: " & Err.Number & " Description: " & Err.Description
    GetRoundJobInfo_byRirekiNumber = JobInfoLocal
    Exit Function
End Function
Public Function GetJobInfo_By_RirekiNumberandKishuInfo(longargRirekiNumber As Long, ByRef argKishuInfo As typKishuInfo) As typJobInfo
    '履歴の連番部分とKishuInfoを貰って、JobInfoを返す
    Dim strTableName As String
    Dim strFieldRirekiNumber As String
    Dim strFieldJobNumber As String
    Dim strFieldInitialDate As String
    Dim strJobNumber As String
    Dim strInitialDate As String
    Dim dicReturn As Dictionary
    Dim JobInfoLocal As typJobInfo
    Dim dbJobinfoByNumber As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Set dbJobinfoByNumber = New clsSQLiteHandle
    Set sqlbC = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    '各フィールド設定
    strTableName = sqlbC.addQuote(Table_JobDataPri & argKishuInfo.KishuName)
    strFieldRirekiNumber = sqlbC.addQuote(Job_RirekiNumber)
    strFieldJobNumber = sqlbC.addQuote(Job_Number)
    strFieldInitialDate = sqlbC.addQuote(Field_Initialdate)
    dbJobinfoByNumber.SQL = "SELECT " & strFieldJobNumber & "," & strFieldInitialDate & " FROM " & strTableName
    dbJobinfoByNumber.SQL = dbJobinfoByNumber.SQL & " WHERE " & strFieldRirekiNumber & " = @RirekiNumber ;"
    Set dbJobinfoByNumber.NamedParm = dbJobinfoByNumber.GetNamedList("@RirekiNumber", Int32, longargRirekiNumber)
    'パラメータバインドありでSQL実行
    dbJobinfoByNumber.Do_SQL_Use_NamedParm_NO_Transaction
    Set dicReturn = dbJobinfoByNumber.RS_Array_Dictionary
    If dbJobinfoByNumber.RecordCount = 0 Then
        '検索失敗してるよ
        Debug.Print "GetJobInfo_By_Rrirekinumber DB Serch result 0"
        GetJobInfo_By_RirekiNumberandKishuInfo = JobInfoLocal
        Set dbJobinfoByNumber = Nothing
        Set sqlbC = Nothing
        Exit Function
    End If
    Set dbJobinfoByNumber = Nothing
    Set sqlbC = Nothing
    strJobNumber = dicReturn("1")(Job_Number)
    strInitialDate = dicReturn("1")(Field_Initialdate)
    '取得したJob番号とInitialDateを元に、JobInfoを取得する
    JobInfoLocal = ReturnJobInfo_by_JobNumber(argKishuInfo, strJobNumber, strInitialDate)
    GetJobInfo_By_RirekiNumberandKishuInfo = JobInfoLocal
    Exit Function
ErrorCatch:
    Debug.Print "GetJobInfo_ByNumber code: " & Err.Number & " Description: " & Err.Description
    GetJobInfo_By_RirekiNumberandKishuInfo = JobInfoLocal
    Exit Function
End Function
Public Function ReturnDivideChrByJobNumber(ByVal strargTableName As String, ByVal strargJobNumber As String, ByVal strargInputDate As String) As Variant
    Dim strSQL As String
    Dim vararrDivideChr As Variant
    Dim dbDivideChr As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Set dbDivideChr = New clsSQLiteHandle
    #If READLOCAL Then
        dbDivideChr.LocalMode = True
    #Else
        dbDivideChr.LocalMode = False
    #End If
    Set sqlbC = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    strSQL = strSQL & "SELECT " & sqlbC.addQuote(Job_KanbanChr) & " AS ""分割文字列"",0 AS ""シート数"",COUNT("
    strSQL = strSQL & sqlbC.addQuote(Job_Rireki) & ") AS ""枚数"", 0 as ""ラック数"",MIN(" & sqlbC.addQuote(Job_Rireki) & ") AS ""スタート履歴"",MAX("
    strSQL = strSQL & sqlbC.addQuote(Job_Rireki) & ") as ""エンド履歴"" FROM " & sqlbC.addQuote(strargTableName)
    strSQL = strSQL & " WHERE " & sqlbC.addQuote(Job_Number) & " = " & sqlbC.addQuote(strargJobNumber) & " AND "
    strSQL = strSQL & sqlbC.addQuote(Field_Initialdate) & " = " & sqlbC.addQuote(strargInputDate) & " AND "
    strSQL = strSQL & sqlbC.addQuote(Job_KanbanChr) & " IS NOT NULL GROUP BY "
    strSQL = strSQL & sqlbC.addQuote(Job_KanbanChr) & "," & sqlbC.addQuote(Job_Number) & "," & sqlbC.addQuote(Field_Initialdate)
    strSQL = strSQL & " ORDER BY " & sqlbC.addQuote(Job_RirekiNumber) & " ASC;"
    dbDivideChr.SQL = strSQL
    dbDivideChr.DoSQL_No_Transaction
    vararrDivideChr = dbDivideChr.RS_Array(boolPlusTytle:=True)
    ReturnDivideChrByJobNumber = vararrDivideChr
    Set dbDivideChr = Nothing
    Set sqlbC = Nothing
    Exit Function
ErrorCatch:
    Set dbDivideChr = Nothing
    Set sqlbC = Nothing
    Debug.Print "ReturnDivideChrByJobNumber code: " & Err.Number & " Description: " & Err.Description
    Exit Function
End Function
Public Function GetNextKanbanChrByTableName(ByVal strargTableName As String) As String
    'JOB関係なく、最大履歴の分割文字列の「次の」使用候補を文字列型で返す
    Dim strLastKanbanChr As String
    Dim dbNextKanbanChr As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Dim strSQL As String
    Dim vararrNextKanban As Variant
    Set dbNextKanbanChr = New clsSQLiteHandle
    #If READLOCAL Then
        dbNextKanbanChr.LocalMode = True
    #Else
        dbNextKanbanChr.LocalMode = False
    #End If
    Set sqlbC = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    '最後の文字列を取得するSQL
    strSQL = strSQL & "SELECT " & sqlbC.addQuote(Job_KanbanChr) & " FROM " & sqlbC.addQuote(strargTableName) & " WHERE " & sqlbC.addQuote(Job_RirekiNumber)
    strSQL = strSQL & " = (SELECT MAX(" & sqlbC.addQuote(Job_RirekiNumber) & ") FROM " & sqlbC.addQuote(strargTableName)
    strSQL = strSQL & " WHERE " & sqlbC.addQuote(Job_KanbanChr) & " IS NOT NULL);"
    dbNextKanbanChr.SQL = strSQL
    dbNextKanbanChr.DoSQL_No_Transaction
    If dbNextKanbanChr.RecordCount = 0 Then
        'レコードカウント0の場合はAを返して終わり
        GetNextKanbanChrByTableName = Chr(MIN_Kanban_ChrCode)
        Exit Function
    End If
    vararrNextKanban = dbNextKanbanChr.RS_Array(boolPlusTytle:=False)
    Set dbNextKanbanChr = Nothing
    Set sqlbC = Nothing
    strLastKanbanChr = UCase(vararrNextKanban(0, 0))
    '定義している最大文字コードを超えるかどうかで処理を分岐する
    If Asc(strLastKanbanChr) + 1 > MAX_Kanban_ChrCode Then
        '超える場合は、最小値（A）を返してやる
        GetNextKanbanChrByTableName = Chr(MIN_Kanban_ChrCode)
        Exit Function
    Else
        '超えない場合は、文字コード+1の文字を返してやる
        GetNextKanbanChrByTableName = Chr(Asc(strLastKanbanChr) + 1)
        Exit Function
    End If
ErrorCatch:
    Set dbNextKanbanChr = Nothing
    Set sqlbC = Nothing
    Debug.Print "GetNextKanbanChrByTableName code: " & Err.Number & " Description: " & Err.Description
    Exit Function
End Function
Public Function GetNextKanbanRirekiByJobNumber(ByVal strargTableName As String, ByVal strargJobNumber As String, ByVal strargInitialDate As String) As String
    'Job内で、KanbanChrがNullのやつの最小履歴（= 看板の次の履歴）を取得する
    Dim dbNextKanbanRireki As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Dim strSQL As String
    Dim vararrNextKanbanRireki As Variant
    Set dbNextKanbanRireki = New clsSQLiteHandle
    #If READLOCAL Then
        dbNextKanbanRireki.LocalMode = True
    #Else
        dbNextKanbanRireki.LocalMode = False
    #End If
    Set sqlbC = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    'SQL組み立て
    strSQL = strSQL & "SELECT MIN(" & sqlbC.addQuote(Job_Rireki) & ") FROM " & sqlbC.addQuote(strargTableName)
    strSQL = strSQL & " WHERE " & sqlbC.addQuote(Job_Number) & " = " & sqlbC.addQuote(strargJobNumber) & " AND "
    strSQL = strSQL & sqlbC.addQuote(Field_Initialdate) & " = " & sqlbC.addQuote(strargInitialDate)
    strSQL = strSQL & " AND " & sqlbC.addQuote(Job_KanbanChr) & " IS NULL;"
    dbNextKanbanRireki.SQL = strSQL
    dbNextKanbanRireki.DoSQL_No_Transaction
    If dbNextKanbanRireki.RecordCount = 0 Then
        '条件に合うものが無かった
        GetNextKanbanRirekiByJobNumber = ""
        Exit Function
    End If
    vararrNextKanbanRireki = dbNextKanbanRireki.RS_Array
    Set dbNextKanbanRireki = Nothing
    Set sqlbC = Nothing
    If IsNull(vararrNextKanbanRireki(0, 0)) Then
        GetNextKanbanRirekiByJobNumber = "残り枚数0なので新規Job分割はできません" & vbCrLf & "次のJobを選択して下さい"
        Exit Function
    End If
    GetNextKanbanRirekiByJobNumber = CStr(vararrNextKanbanRireki(0, 0))
    Exit Function
ErrorCatch:
    Set dbNextKanbanRireki = Nothing
    Set sqlbC = Nothing
    Debug.Print "GetNextKanbanRireki code: " & Err.Number & " Description: " & Err.Description
End Function
Public Function UpdateKanbanChrByJobNumberMaisuu(strargTableName As String, strargKanbanChr As String, strargStartRireki As String, longargMaisuu As Long, argKishuInfo As typKishuInfo) As Boolean
    '与えられた条件を元に看板データのUpdateを行う
    Dim intRackTotal As Integer
    Dim intCurrentRack As Integer
    Dim longCurrentMaisuuStart As Long
    Dim longCurrentMaisuuEnd As Long
    Dim longStartRirekiNumber As Long
    Dim longEndRirekiNumber As Long
    Dim longMaiPerRack As Long
    Dim strSQL As String
    Dim dbUpdateKanban As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Dim isCollect As Boolean
    On Error GoTo ErrorCatch
    If longargMaisuu <= 0 Then
        MsgBox "枚数に0以下が指定されました。処理を中止します"
        UpdateKanbanChrByJobNumberMaisuu = False
        Exit Function
    End If
    'ラックあたりの枚数を求める
    longMaiPerRack = CLng(argKishuInfo.MaiPerSheet) * CLng(argKishuInfo.SheetPerRack)
    'トータルのラック数を求める
    intRackTotal = Application.WorksheetFunction.RoundUp(CDbl(longargMaisuu) / CDbl(longMaiPerRack), 0)
    'スタート履歴と終了履歴を求める
    longStartRirekiNumber = CLng(Right(strargStartRireki, argKishuInfo.RenbanKetasuu))
    longEndRirekiNumber = longStartRirekiNumber + longargMaisuu - 1
    intCurrentRack = 1
    Do While intCurrentRack <= intRackTotal
    '最終ラックかどうかで処理を分岐する
        Select Case intCurrentRack = intRackTotal
        Case True
            '最終ラックの場合
            longCurrentMaisuuStart = longStartRirekiNumber + ((intCurrentRack - 1) * longMaiPerRack)
            longCurrentMaisuuEnd = longEndRirekiNumber
        Case False
            '途中のラックの場合
            longCurrentMaisuuStart = longStartRirekiNumber + ((intCurrentRack - 1) * longMaiPerRack)
            longCurrentMaisuuEnd = longCurrentMaisuuStart + longMaiPerRack - 1
        End Select
        'SQL生成、処理、トランザクション処理ありで
        'ここでDBとSQLBuliderのインスタンス生成
        Set dbUpdateKanban = New clsSQLiteHandle
        #If READLOCAL Then
            dbUpdateKanban.LocalMode = True
        #Else
            dbUpdateKanban.LocalMode = False
        #End If
        Set sqlbC = New clsSQLStringBuilder
        strSQL = ""
        strSQL = strSQL & "UPDATE " & sqlbC.addQuote(strargTableName) & " SET " & sqlbC.addQuote(Job_KanbanChr) & " = " & sqlbC.addQuote(strargKanbanChr)
        strSQL = strSQL & "," & sqlbC.addQuote(Job_KanbanNumber) & " = " & intCurrentRack
        strSQL = strSQL & " WHERE " & sqlbC.addQuote(Job_RirekiNumber) & " BETWEEN " & longCurrentMaisuuStart & " AND " & longCurrentMaisuuEnd & ";"
        dbUpdateKanban.SQL = strSQL
        isCollect = dbUpdateKanban.Do_SQL_With_Transaction
        Set dbUpdateKanban = Nothing
        Set sqlbC = Nothing
        If Not isCollect Then
            MsgBox "看板情報アップデート中にエラー発生"
            Debug.Print "UpdateKanbanChrByJobMaisuu Table:" & strargTableName & " StarNumber: " & longCurrentMaisuuStart & " EndNumber: " & longCurrentMaisuuEnd
            Exit Function
        End If
        '1ラック目の処理完了
        'ラックカウントインクリメント
        intCurrentRack = intCurrentRack + 1
     Loop
     UpdateKanbanChrByJobNumberMaisuu = True
     Exit Function
ErrorCatch:
     Debug.Print "UpdateKanbanChrByJobNumberMaisuu code: " & Err.Number & " Description: " & Err.Description
End Function
Public Function GetOriginalDBSchemaByKishuName(ByVal strargKishuName As String) As Dictionary
    '与えられた機種名(Job_Kishuname)より、テーブルとインデックスのスキーマ(sql）をDictionaryで返す
    Dim dbOrigin As clsSQLiteHandle
    Dim dicLocalSchema As Dictionary
    On Error GoTo ErrorCatch
    Set dbOrigin = New clsSQLiteHandle
    Set dicLocalSchema = New Dictionary
    dbOrigin.SQL = "SELECT type,name,sql FROM sqlite_schema WHERE sql IS NOT NULL AND "
    dbOrigin.SQL = dbOrigin.SQL & "name LIKE ""%" & Table_JobDataPri & strargKishuName & "%"";"
    dbOrigin.DoSQL_No_Transaction
    Set dicLocalSchema = dbOrigin.RS_Array_Dictionary
    Set GetOriginalDBSchemaByKishuName = dicLocalSchema
    Set dbOrigin = Nothing
    Set dicLocalSchema = Nothing
    Exit Function
ErrorCatch:
    Debug.Print "GetOriginalDBSchemaByKishuName code: " & Err.Number & " Description: " & Err.Description
    Set dbOrigin = Nothing
    Set dicLocalSchema = Nothing
    Exit Function
End Function
Public Function CopyDBTableRemote_To_Local(ByVal strargTableName As String, ByVal strargLocalDBFilePath As String, Optional ByVal strargRemoteDbFilePath As String, _
                                            Optional ByVal longargLastNumberArea As Long) As Boolean
    'リモートをオリジナルとして、ローカルに選択コピーする
    'longargLastNunmberAreaで、最新〇件に絞って抽出するようにする、していない場合は最初から最後まで（時間かかるよ・・・)
    Dim strRemoteDBPath As String
    Dim dbCopyTable As clsSQLiteHandle
    Dim strRmoteTableName As String
    Dim strLocalTableName As String
    Dim strSrcWhereField As String
    Dim sqlbC As clsSQLStringBuilder
    Dim strSQL As String
    Dim isCollect As Boolean
    Dim longSrcStartRirekiNumber As Long            'コピー元テーブルの最小履歴番号（連番部分のみ）
    Dim longSrcEndRirekiNumber As Long              'コピー元テーブルの最大履歴番号
    Set dbCopyTable = New clsSQLiteHandle
    Set sqlbC = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    'まずはリモートDBファイルパスを決定する
    If strargRemoteDbFilePath = "" Then
        strRemoteDBPath = constDatabasePath & "\" & constJobNumberDBname
    Else
        strRemoteDBPath = strargRemoteDbFilePath
    End If
    '次にリモートとローカルのテーブル名
    strRmoteTableName = """remote""." & sqlbC.addQuote(strargTableName)
    strLocalTableName = sqlbC.addQuote(strargTableName)
    '次にリモートDBをアタッチする
    strSQL = "ATTACH " & sqlbC.addQuote(strRemoteDBPath) & " AS ""remote"";"
    dbCopyTable.SQL = strSQL
    dbCopyTable.DBPath = strargLocalDBFilePath
    isCollect = dbCopyTable.DoSQL_No_Transaction()
    If Not isCollect Then
        MsgBox "リモートDBアタッチ時にエラー発生"
        GoTo ErrorCatch
        Exit Function
    End If
    '次に、リモートからローカルへのコピー処理実行（ここが重い！）
    'トランザクション有でSQL実行
    'パラメータバインドを使用する
    MsgBox "まだ作成途中だよ_CopyDBTableRemote_TO_Local"
    isCollect = dbCopyTable.DoSQL_No_Transaction("BEGIN TRANSACTION")
    If Not isCollect Then
        MsgBox "トランザクション処理開始失敗、処理を中断します"
        CopyDBTableRemote_To_Local = False
        dbCopyTable.RollBackTransaction
        GoTo ErrorCatch
        Exit Function
    End If
    'INSERT INTO T_JobData_Test15 SELECT * FROM (SELECT * FROM remote.T_JobData_Test15
    'WHERE RirekiNumber BETWEEN 1210001 and 2400000) srcT WHERE NOT EXISTS
    ' (SELECT * FROM T_JobData_Test15 dstT WHERE srcT.Rireki = dstT.Rireki);
    strSrcWhereField = sqlbC.addQuote(Job_RirekiNumber)
    strSQL = "INSERT INTO " & strLocalTableName & " SELECT * FROM (SELECT * FROM " & strRemoteDBPath
    strSQL = strSQL & " WHERE " & strSrcWhereField & " BETWEEN @SrcRirekiNumberStart and @SrcRirekiNumberEnd) SrcT WHERE NOT EXISTS"
    strSQL = strSQL & " (SELECT * FROM " & strLocalTableName & " DstT WHERE @SrcRireki = @DstRireki);"
    dbCopyTable.SQL = strSQL
    Set dbCopyTable.NamedParm = dbCopyTable.GetNamedList("@SrcRirekiNumberStart", Int32, Job_RirekiNumber)
    If Not isCollect Then
        MsgBox "リモートDBからのコピー時にエラー発生"
        GoTo ErrorCatch
        Exit Function
    End If
    isCollect = dbCopyTable.DoSQL_No_Transaction("COMMIT TRANSACTION")
    Set dbCopyTable = Nothing
    Set sqlbC = Nothing
    CopyDBTableRemote_To_Local = True
    Exit Function
ErrorCatch:
    Debug.Print "CopyDBTableRemote_To_Local code: " & Err.Number & " Description: " & Err.Description
    Set dbCopyTable = Nothing
    Set sqlbC = Nothing
    CopyDBTableRemote_To_Local = False
    Exit Function
End Function
Public Function CountKanbanChr(ByVal strargJobNumber As String, ByVal strargInitialInputDate As String) As Integer
    'Job番号基準で、分割文字列の数を返します
End Function
Public Function UpdateLastRirekNumber_atKishuTable(ByVal strargKishuName, ByVal longargLastRirekiNumber As Long, ByVal longargLastNumberField As LastRirekiNumber, _
                                                    Optional ByVal boolUseNewNumberOnly As Boolean = False) As Boolean
    '機種テーブルの最終履歴を更新までしちゃう
    '最後のUserNewNumberOnlyにTrueをセットすると、常に与えられた番号で上書きする（主にメンテナンス用）
    Dim strSQL As String
    Dim strTargetTable As String
    Dim strUpdateField As String
    Dim strWhereField As String
    Dim dbUpdateLastRirekiNumber As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Dim isCollect As Boolean
    Set dbUpdateLastRirekiNumber = New clsSQLiteHandle
    Set sqlbC = New clsSQLStringBuilder
    strTargetTable = sqlbC.addQuote(Table_Kishu)
    strWhereField = sqlbC.addQuote(Kishu_KishuName)
    'パラメータバインド用リスト設定
    Set dbUpdateLastRirekiNumber.NamedParm = dbUpdateLastRirekiNumber.GetNamedList("@NewLastNumber", Int32, longargLastRirekiNumber)
    Set dbUpdateLastRirekiNumber.NamedParm = dbUpdateLastRirekiNumber.GetNamedList("@WhereCondition", Text, strargKishuName)
    'Job番号か看板番号かでフィールドを決定する
    Select Case longargLastNumberField
    Case LastRirekiNumber.JobNumberField
        'Job番号のラストを更新したい
        strUpdateField = sqlbC.addQuote(Kishu_Jobnumber_Lastnumber)
    Case LastRirekiNumber.KanbanChrField
        '看板のラストを更新したい
        strUpdateField = sqlbC.addQuote(Kishu_Kanbanchr_Lastnumber)
    Case Else
        MsgBox "指定外の列定義が選択されました"
        UpdateLastRirekNumber_atKishuTable = False
        Exit Function
    End Select
    strSQL = "UPDATE " & strTargetTable & " SET " & strUpdateField
    If boolUseNewNumberOnly Then
        'メンテナンスモード、常に引数で与えられた数で更新しちゃう
        strSQL = strSQL & " = @NewLastNumber WHERE "
    Else
        '通常モード、機種テーブルと最新数値で、値の大きいほうを残す
        strSQL = strSQL & " = CASE WHEN @NewLastNumber <= " & strUpdateField & " THEN " & strUpdateField & " ELSE @NewLastNumber END WHERE "
    End If
    strSQL = strSQL & strWhereField & " = @WhereCondition;"
    dbUpdateLastRirekiNumber.SQL = strSQL
    isCollect = dbUpdateLastRirekiNumber.Do_SQL_Use_NamedParm_NO_Transaction
    If Not isCollect Then
        Debug.Print "UpdateLastRirekiNumber_asKishuTable fail"
        UpdateLastRirekNumber_atKishuTable = False
        Exit Function
    End If
    UpdateLastRirekNumber_atKishuTable = True
    Exit Function
End Function
Private Function GetMinimumRirekiNumber_byTableName(ByVal strargTableName As String) As Long
    '与えられたテーブルの連番最小値を返す
    Dim strFieldNumber
    Dim varReturn As Variant
    Dim dbMinNumber As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Set dbMinNumber = New clsSQLiteHandle
    Set sqlbC = New clsSQLStringBuilder
    strFieldNumber = sqlbC.addQuote(Job_RirekiNumber)
    dbMinNumber.SQL = "SELECT MIN(" & strFieldNumber & ") FROM " & strargTableName & " ;"
    dbMinNumber.DoSQL_No_Transaction
    varReturn = dbMinNumber.RS_Array
    Set dbMinNumber = Nothing
    Set sqlbC = Nothing
    If Not IsNumeric(varReturn(0, 0)) Then
        '取得失敗してるよね・・
        Debug.Print "GetMinNumber fail"
        GetMinimumRirekiNumber_byTableName = 0
        Exit Function
    Else
        GetMinimumRirekiNumber_byTableName = CLng(varReturn(0, 0))
        Exit Function
    End If
End Function
Public Function GetLastRirekiNumber_byKishuTable(ByVal strargKishuName, ByVal longargLastNumberField As LastRirekiNumber, Optional ByVal boolRenumberIfZero) As Long
    '機種テーブルの最終履歴を取得する
    '引数で指定されたLastRirekiNumber Enumの値により処理を分岐（看板かJob番号か）
    'RenumberIfZeroがTrueの場合、結果が0だった場合に、PublicModule.RenumberKishuTableLastNumberを呼ぶ
    Dim varResult As Variant
    Dim strSelectField As String                'selectのフィールド名、これはLastRirekiNumberの値により分岐
    Dim strWherField As String                  'whereの検索フィールド名
    Dim isCollect As Boolean
    Dim dbLastNumber As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Set dbLastNumber = New clsSQLiteHandle
    Set sqlbC = New clsSQLStringBuilder
    'パラメータバインド使用
    '                   SELECT JobNumber_LastNumber FROM T_Kishu WHERE KishuName = "Test15";
    strWherField = sqlbC.addQuote(Kishu_KishuName)
    Set dbLastNumber.NamedParm = dbLastNumber.GetNamedList("@WhereCondition", Text, strargKishuName)
    Select Case longargLastNumberField
    Case LastRirekiNumber.JobNumberField
        'JobNumberのラストが欲しい
        strSelectField = sqlbC.addQuote(Kishu_Jobnumber_Lastnumber)
    Case LastRirekiNumber.KanbanChrField
        '看板のラストが欲しい
        strSelectField = sqlbC.addQuote(Kishu_Kanbanchr_Lastnumber)
    Case Else
        MsgBox "指定外の数値がセットされました。処理を中断します"
        Debug.Print "GetLastRIrekiNumber_byKishuTable Unknown Field Nuber"
        GetLastRirekiNumber_byKishuTable = -1
        Exit Function
    End Select
    dbLastNumber.SQL = "SELECT " & strSelectField & " FROM T_Kishu WHERE " & strWherField & " = @WhereCondition;"
    isCollect = dbLastNumber.Do_SQL_Use_NamedParm_NO_Transaction
    If Not isCollect Then
        Debug.Print "GetLastRIrekiNumber_byKishuTable fail"
        GetLastRirekiNumber_byKishuTable = 0
        Set dbLastNumber = Nothing
        Set sqlbC = Nothing
        Exit Function
    End If
    varResult = dbLastNumber.RS_Array
    If dbLastNumber.RecordCount = 0 Then
        Debug.Print "GetLastRirekiNumber_BykishuTable 条件に一致するデータなし（Nullの可能性）"
        GetLastRirekiNumber_byKishuTable = 0
        Set dbLastNumber = Nothing
        Exit Function
    End If
    Set dbLastNumber = Nothing
    Set sqlbC = Nothing
    If IsNull(varResult(0, 0)) Or (varResult(0, 0)) = 0 Then
        Debug.Print "GetLastRirekiNumber Result is Null or zero"
        If boolRenumberIfZero Then
            'ゼロで再取得指示がある場合
            Call RenumberKishuTableLastNumber
            '再帰でもう一度呼ぶ、ただし再取得なしで
            GetLastRirekiNumber_byKishuTable = GetLastRirekiNumber_byKishuTable(strargKishuName, longargLastNumberField, boolRenumberIfZero:=False)
        Else
            '通常（何もしない）
            GetLastRirekiNumber_byKishuTable = 0
        End If
        Exit Function
    Else
        GetLastRirekiNumber_byKishuTable = CLng(varResult(0, 0))
    End If
    Exit Function
End Function