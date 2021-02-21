Attribute VB_Name = "SQL_View_To"
Option Explicit
'将来的にViewテーブルにしたいSQLを集めます
Public Function ReturnJobNumber_For_KanbanDivide(ByVal strargTableName As String) As Variant
    Dim strSQL As String
    Dim vararrKanbanJobNumber As Variant
    Dim dbKanbanJobNumber As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Set dbKanbanJobNumber = New clsSQLiteHandle
    Set sqlbC = New clsSQLStringBuilder
    On Error GoTo ErrorCatch
    strSQL = strSQL & "SELECT " & sqlbC.addQuote(Job_Number) & " as ""Job番号"", " & sqlbC.addQuote(Field_Initialdate) & " as ""登録日時"",count(*) - count("
    strSQL = strSQL & sqlbC.addQuote(Job_KanbanChr) & ") as ""残り枚数"" FROM " & sqlbC.addQuote(strargTableName)
'    strSQL = strSQL & " WHERE " & sqlbC.addQuote(Job_KanbanChr) & " IS NULL GROUP BY " & sqlbC.addQuote(Job_Number) & "," & sqlbC.addQuote(Field_Initialdate)
    strSQL = strSQL & " GROUP BY " & sqlbC.addQuote(Job_Number) & "," & sqlbC.addQuote(Field_Initialdate)
    strSQL = strSQL & " ORDER BY " & sqlbC.addQuote(Job_RirekiNumber) & " ASC;"
    dbKanbanJobNumber.SQL = strSQL
    dbKanbanJobNumber.DoSQL_No_Transaction
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
Public Function ReturnDivideChrByJobNumber(ByVal strargTableName As String, ByVal strargJobNumber As String, ByVal strargInputDate As String) As Variant
    Dim strSQL As String
    Dim vararrDivideChr As Variant
    Dim dbDivideChr As clsSQLiteHandle
    Dim sqlbC As clsSQLStringBuilder
    Set dbDivideChr = New clsSQLiteHandle
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