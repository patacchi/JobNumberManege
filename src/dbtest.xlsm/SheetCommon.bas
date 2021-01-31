Attribute VB_Name = "SheetCommon"
Public Sub SheetToDB()
    'Job番号が変わるか、履歴が連続じゃなくないか、機種が変わった時を一区切りとしてSQL発行する
    Dim rngJobNumberStart As Range
    Dim rngRirekiStart As Range
    Dim longLastRow As Long
    Dim longCurrentRow As Long
    Dim longcurrentMaisuu As Long
    Dim longSQLStartRireki As Long
    Dim sqlbSheetToDB As clsSQLStringBuilder
    Set sqlbSheetToDB = New clsSQLStringBuilder
    Dim varRirekiNumber As Variant
    Dim varJobNumber As Variant
    Dim longStartRirekiNumber As Long
    Dim longEndRirekiNumber As Long
    Dim longDuplicateRirekiNumber As Long
    Dim dblTimer As Double
    Dim boolArrowNextLoop As Boolean
    Dim boolSameData As Boolean
    Dim boolSameKishu As Boolean
    Dim KishuInfo As typKishuInfo
    Dim newKishuInfo As typKishuInfo        '新旧のkishuinfoを比較する必要があるため
    Dim isCollect As Boolean
    On Error GoTo ErrorCatch
    Set rngJobNumberStart = Application.InputBox(prompt:="ジョブ番号の最初のセルを選択して下さい", Type:=8)
    Set rngRirekiStart = Application.InputBox(prompt:="履歴番号の最初のセルを選択して下さい", Type:=8)
    'シート内機種混在対応のため、機種テーブルの全情報を配列で受け取るにゃん
    'とりあえず最終行を取得して、シートの値を配列に格納する
    '1ベースの配列なので注意
    'array(2,1) みたいな感じで取り出す
    longLastRow = Cells(Rows.Count, rngJobNumberStart.Column).End(xlUp).Row
    varJobNumber = Range(rngJobNumberStart, rngJobNumberStart.offset(longLastRow - rngJobNumberStart.Row)).Value
    varRirekiNumber = Range(rngRirekiStart, rngRirekiStart.offset(longLastRow - rngRirekiStart.Row)).Value
    'Do Whileでループして、currentMaisuuをインクリメントしていく（CurrentRowも当然）
    'で､ジョブ番号か履歴の増分が変化して時点でいったんboolArrowNextLoopにFalseセット、SQL発行､currentMaisuuだけは1にもどしてやる
    longcurrentMaisuu = 0
    longCurrentRow = 1
    boolArrowNextLoop = False
    'イニシャルKishuInfo
    '機種混在してるので、毎回確認する必要があり
    KishuInfo = getKishuInfoByRireki(CStr(varRirekiNumber(1, 1)))
    'ここで機種登録成功フラグ立ってなかったら即終了で
    If Not boolRegistOK Then
        Debug.Print "機種登録フラグNGにより終了"
        GoTo ErrorCatch
        Exit Sub
    End If
    newKishuInfo = KishuInfo
    boolArrowNextLoop = True
    boolSameData = True
    boolSameKishu = True
    Debug.Print "Sheet_to_DB時間計測開始"
    dblTimer = timer
    Do
        Do
            '最初に枚数をインクリメント
            longcurrentMaisuu = longcurrentMaisuu + 1
            'ここで次の行が連続データならSameDataフラグをTrueに
            '次の行が範囲内ならArrowNextLoopフラグをTrueに
            '次の行が存在するかもチェックしないと・・・
            '次の行のKishuInfo取って、KishuHeaderで比較、違うならいったんそこで区切り
            If Not longCurrentRow + 1 <= UBound(varRirekiNumber) Then
                boolArrowNextLoop = False
            ElseIf CLng(Right(varRirekiNumber(longCurrentRow, 1), KishuInfo.RenbanKetasuu)) = _
                CLng(Right(varRirekiNumber(longCurrentRow + 1, 1), KishuInfo.RenbanKetasuu) - 1) And _
                varJobNumber(longCurrentRow, 1) = varJobNumber(longCurrentRow + 1, 1) Then
                boolSameData = True
            Else
                boolSameData = False
            End If
            If longCurrentRow + 1 <= UBound(varRirekiNumber) Then
                '次も行けるならnewKishuInfo取ってみる
                '旧機種のKishuInfoで次のヘッダ取っちゃえばいいのでは・・・
                'ちゃんと毎回kishuinfo取らないとだめ、数値の桁数が機種によりけり
                newKishuInfo = getKishuInfoByRireki(CStr(varRirekiNumber(longCurrentRow + 1, 1)))
                If KishuInfo.KishuHeader = newKishuInfo.KishuHeader Then
                    'KiahuInfoのKishuHeaderが同じなので、機種も同じでしょう
                    boolSameKishu = True
                Else
                    boolSameKishu = False
                End If
            End If
            'longcurrentRowをインクリメントし、配列上限を超えるようだったらSQL実行したのちに処理を終了する
            longCurrentRow = longCurrentRow + 1
            If longCurrentRow > UBound(varRirekiNumber) Then
                Exit Do
            End If
        Loop While boolArrowNextLoop And boolSameData And boolSameKishu
        'ここでcurrentMaisuu分のSQLを発行する
        longSQLStartRireki = longCurrentRow - longcurrentMaisuu
        With sqlbSheetToDB
'            .BulkCount = 800
            .StartRireki = CStr(varRirekiNumber(longSQLStartRireki, 1))
            .FieldArray = arrFieldList_JobData
            .JobNumber = CStr(varJobNumber(longSQLStartRireki, 1))
            .Maisu = longcurrentMaisuu
            .RenbanKeta = KishuInfo.RenbanKetasuu
            .TableName = Table_JobDataPri & KishuInfo.KishuName
        End With
        '重複データのチェック
        longStartRirekiNumber = CLng(Right(sqlbSheetToDB.StartRireki, KishuInfo.RenbanKetasuu))
        longEndRirekiNumber = longStartRirekiNumber + sqlbSheetToDB.Maisu - 1
        longDuplicateRirekiNumber = GetRecordCountSimple(sqlbSheetToDB.TableName, Job_RirekiNumber, _
                                    "BETWEEN " & longStartRirekiNumber & " AND " & longEndRirekiNumber & ";")
        If longDuplicateRirekiNumber >= 1 Then
            '重複ありなので、シートからの登録は無視する、最初の履歴位は表示してやろうか
            MsgBox sqlbSheetToDB.StartRireki & " から始まる履歴で、 " & longDuplicateRirekiNumber & " 件の重複があったようです。今回の履歴は処理をスキップします。"
        Else
            Call sqlbSheetToDB.CreateInsertSQL
        End If
        '次の機種用に、さっき取っておいた次の行分のKishuInfoに差し替え
        'KishuInfoは同機種だと初期化の時に同じものが入ってるはず
        KishuInfo = newKishuInfo
        longcurrentMaisuu = 0
    Loop While longCurrentRow < UBound(varRirekiNumber)
    MsgBox "シートからの変換完了 " & UBound(varRirekiNumber) & " 件のデータを " & timer - dblTimer & "秒で処理しました"
    Debug.Print "シートからの変換完了 " & UBound(varRirekiNumber) & " 件のデータを " & timer - dblTimer & "秒で処理しました"
    Set sqlbSheetToDB = Nothing
    Exit Sub
ErrorCatch:
        Debug.Print "btnSheet to DB Code: " & Err.Number & " Description: " & Err.Description
        Exit Sub
End Sub