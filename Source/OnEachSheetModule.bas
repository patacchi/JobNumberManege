Attribute VB_Name = "OnEachSheetModule"
Option Explicit
'Option Base 1

Public Function Conversion_Target_To_Array(ByRef rangeTargetArg As Range, ByRef vararrCreate As Variant) As Boolean
    'Worksheet_Select_Changeイベント発生時に、グローバル変数に開始・終了行数、値の配列を退避させるプログラム
    
    Dim timeStart As Double
    '時間計測
    timeStart = timer()
    'Selectionの開始・最終行格納
    longRowStart = rangeTargetArg.Row
    If longRowStart <= constSheetRowStart Then
        longRowStart = constSheetRowStart
    End If
    longRowEnd = rangeTargetArg.Row + rangeTargetArg.Rows.Count - 1
    If longRowEnd < longRowStart Then
        Conversion_Target_To_Array = False
        Exit Function
    End If
    vararrCreate = Range(Cells(longRowStart, constDataStartColumn), Cells(longRowEnd, constDataEndColumn))
    Debug.Print "処理行数は" & longRowEnd - longRowStart + 1 & vbCrLf & "処理時間は" & timer() - timeStart
    Conversion_Target_To_Array = True

End Function

Public Sub Btn_Add_All()
    '全履歴追加（Todo）
'    Dim strDatabaseName As String
'    Dim strConnectionString As String
'    Dim bytRetCode As Byte
'    Dim fso As New Scripting.FileSystemObject
'    If Application.EnableEvents = False Then
'        Application.EnableEvents = True
'    End If
'    'カレントディレクトリの移動と、データベース名の取得
'    strDatabaseName = PublicModule.ChcurrentAndReturnDBName
'    'データベースファイルの存在有無チェック
'    If fso.FileExists(strDatabaseName) = True Then
'        'あった場合
'        'データ追加処理（Todo)
'
'    Else
'        '無かった場合
'        MsgBox "データベースファイルが無いので新規作成します"
'        '接続文字列の取得
'        strConnectionString = PublicModule.GetConnectionString(strDatabaseName)
'        '空データベースの作成
'        'bytRetCode = PublicModule.CreateMDB(strDatabaseName, strConnectionString)
'        MsgBox "空のデータベースファイル作成が完了しました"
'        '新規作成後、データ追加処理（Todo)
'    End If
'    Set fso = Nothing
End Sub

Public Function Backup_From_To_Data(ByRef rngTarget As Range, ByRef typMaisu() As typMaisuuRireki)
    Dim intProceccCount As Integer  'forループ用カウンター
    Dim nameFrom As Name
    Dim nameTo As Name
    Dim bytFromColumn As Byte   'Fromの列番号
    Dim bytToColumn As Byte     'Toの列番号
    'Dim typMaisuLocal() As New typMaisuuRireki
    
    Range(rngTarget.Item(1).Address).Activate
    '各名前定義の列番号取得
    Set nameFrom = GetNameRange(constRirekiFromLabel)
    Set nameTo = GetNameRange(constRirekiToLabel)
    
    bytFromColumn = nameFrom.RefersToRange.Column
    bytToColumn = nameTo.RefersToRange.Column
    
    'typMaisuuRireki型にデータ格納
    For intProceccCount = 1 To rngTarget.Count
        
    Next intProceccCount

End Function

Public Function On_WorkSheet_Change(ByRef rngNewTarget As Range, ByRef rngOldTarget As Range)
    'ここの処理は大幅に見直し必要（Todo）
    'ワークシートにデータチェンジ発生した場合
    '名前の定義にMaisuu_Rageがあるか調べて、ある場合は変更がその範囲か調べて
    '入力された数字に基づいてデータベースのUpdate(Insert)を発行
    'データベース追加の際は、予め入力するデータを配列に格納して参照渡し→追い出しました
    '名前が定義されて無い場合は警告出して処理中断
    '正常に処理終了した場合はセルの色をSelection.Interior.ColorIndex = 35 する（うすいぐりーん）
    
    Dim strSQLString() As String
    Dim intErrCode
    Dim rngName As Name
    Dim intTargetRowCount As Integer
    Dim intAddQty As Integer
    Dim strDatabaseName As String
'
'    Set rngName = GetNameRange(constMaisuu_RangeLabel)
'
'    If rngName Is Nothing Then
'        '枚数の名前定義見つからなかったので終了
'        MsgBox ActiveSheet.Name & "シートの" & constMaisuu_RangeLabel & "名前定義が見つからなかったので処理を中止します"
'        Exit Function
'    End If
'
'    If Intersect(rngNewTarget, rngName.RefersToRange) Is Nothing Then
'        '今回はまいすーの場所じゃなかったのでスルー
'        Exit Function
'    End If
'
'    intTargetRowCount = 0
'    For intTargetRowCount = 1 To rngNewTarget.Count
'        '変更範囲の行数文For回し、各行毎にInsert文の塊をもらう
'        intAddQty = rngNewTarget.Item(intTargetRowCount).Value
'        If intAddQty = 0 Or intAddQty = Empty Then
'            GoTo SkipLoop1
'        End If
'        ReDim strSQLString(intAddQty)
'        strSQLString = PublicModule.CreateInsertSQL(rngNewTarget.Item(intTargetRowCount))
'
'        'If strSQLString = CStr(errcxlDataNothing) Then
'         '   GoTo SkipLoop1
'        'End If
'
'        'データベースに追加する
'        'カレントの移動とデータベース名取得
'        strDatabaseName = PublicModule.ChcurrentAndReturnDBName
'        intErrCode = PublicModule.Mdb_RirekiAdd(strDatabaseName, strSQLString)
'        If intErrCode = 0 Then
'            ActiveCell.Interior.ColorIndex = 35
'        End If
'SkipLoop1:
'        Erase strSQLString
'    Next intTargetRowCount
'    Application.StatusBar = ""
'
'    Select Case intErrCode
'    Case 0
'        '正常終了した場合
'        On_WorkSheet_Change = 0
'    Case 4
'        'データベースファイル見つからなかった場合
'        On_WorkSheet_Change = 4
'    End Select
'
End Function

