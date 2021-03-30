VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKanban 
   Caption         =   "看板分割処理フォーム"
   ClientHeight    =   8475.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15690
   OleObjectBlob   =   "frmKanban.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmKanban"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#const ReadLocal =1 とすると、機種名特定できた時点でローカルに一時テーブルを作成し、そこから検索するようにする
'他でも参照するっぽいので、Vbaprojectのプロパティでグローバル変数として定義
#If READLOCAL = 1 Then
    Private dbLocal As clsSQLiteHandle
    Private strLocalPath
    Private strLocalDBFilePath
#End If
Private Sub btnDeleteLastKanban_Click()
    If listBoxExistingChr.ColumnCount <= 1 Then
        MsgBox "JOB分割履歴が存在しないようなので、削除処理を中止します"
        Exit Sub
    End If
End Sub
Private Sub btnDoDivide_Click()
    Dim KishuLocal As typKishuInfo
    Dim isCollect As Boolean
    Dim longOldJobNumberIndex As Long
    Dim longOldNicknameIndex As Long
    On Error GoTo ErrorCatch
    KishuLocal = GetKishuinfoByNickName(lblNowKishuNickName.Caption)
    If txtBoxNewMaisuu.Text = "" Or txtboxNewSheetQty.Text = "" Or cmbBoxKanbanChr.Text = "" Then
        MsgBox "必要な項目が入力されていないようです。"
        Exit Sub
    End If
    If CLng(txtBoxNewMaisuu.Text) > CLng(lblRemainMaisuu.Caption) Then
        MsgBox "残り枚数を超えています。"
        Exit Sub
    End If
    If CLng(txtBoxNewMaisuu.Text) <= 0 Then
        MsgBox "枚数に0以下がセットされたため、終了します"
        Exit Sub
    End If
    isCollect = UpdateKanbanChrByJobNumberMaisuu(Table_JobDataPri & KishuLocal.KishuName, cmbBoxKanbanChr.Text, lblNextRireki.Caption, txtBoxNewMaisuu, KishuLocal)
    If Not isCollect Then
        MsgBox "看板データの設定時にエラーが発生したようです"
        Exit Sub
    End If
    '終了したのでお掃除
'    strOldJobNumber = cmbBoxJobNumber.List(cmbBoxJobNumber.ListIndex, 0)
    longOldJobNumberIndex = cmbBoxJobNumber.ListIndex
    longOldNicknameIndex = cmbBoxKishuNickName.ListIndex
    Call Clear_Exclude_KishuNickName(boolExcludeJobNumber:=True)
    'リスト再表示
    cmbBoxKishuNickName.ListIndex = -1
    Sleep 20
    cmbBoxJobNumber.ListIndex = -1
    Sleep 20
    cmbBoxKishuNickName.ListIndex = longOldNicknameIndex
    Sleep 20
    cmbBoxJobNumber.ListIndex = longOldJobNumberIndex
    Sleep 100
    listBoxExistingChr.Width = 520
    Sleep 50
    listBoxExistingChr.Width = 525.5
    Sleep 30
    MsgBox "Job分割完了"
    Exit Sub
ErrorCatch:
    Debug.Print "btnDoDivide code: " & Err.Number & " Descriptoin: " & Err.Description
    Exit Sub
End Sub
Private Sub btnPrintKanban_Click()
    If listBoxExistingChr.ListIndex = -1 Or listBoxExistingChr.ListIndex = 0 Then
        MsgBox "看板を作成したい分割番号を選んでからクリックして下さい。"
        Exit Sub
    End If
    '看板作成開始
End Sub
Private Sub btnQRRead_Click()
    Dim KishuLocal As typKishuInfo
    Dim qrLocal As typQRDataField
    On Error GoTo ErrorCatch
    QRField = qrLocal
    frmQRAnalyze.Show
    If QRField.Zuban = "" Then
        Exit Sub
    End If
    KishuLocal = GetKishuinfoByKishuName(QRField.Zuban)
    If KishuLocal.KishuNickName = "" Then
        MsgBox "読み込まれたQRコードの機種情報が見つかりませんでした。Job登録画面から登録して下さい。"
        Exit Sub
    End If
    '機種名コンボボックスにセットしてやる
    cmbBoxKishuNickName.Text = KishuLocal.KishuNickName
    cmbBoxJobNumber.Text = QRField.JobNumber
ErrorCatch:
    Debug.Print "btnQRRead code:" & Err.Number & " Descriptoin: "; Err.Description
End Sub
Private Sub cmbBoxKishuNickName_Change()
    Dim KishuNickName As typKishuInfo
    Dim vararrJobData As Variant
    Dim strListColumnWidth As String
    '違うのを選択したパターンのために、入力変化したら、他の項目を初期化してやる
    Call Clear_Exclude_KishuNickName
    '機種通称名からKishuInfoを引っ張ってくる
    KishuNickName = GetKishuinfoByNickName(cmbBoxKishuNickName.Text)
    '看板一覧表取得SQL大幅改善につき、ローカルコピーは不要になりました
    #If READLOCAL = 1 Then
        'ローカルコピー時は、ここで機種関連テーブルをローカルに取り込む
        Dim dicOriginTableSchema As Dictionary
        Dim varKeyLocal As Variant
        Dim isCollect As Boolean
        Set dbLocal = New clsSQLiteHandle
        Set dicOriginTableSchema = GetOriginalDBSchemaByKishuName(KishuNickName.KishuName)
        'まずはテーブル定義のみを済ませる
        For Each varKeyLocal In dicOriginTableSchema
            'テーブルが存在しない場合のみ作成
            If Not IsTableExist(dicOriginTableSchema(varKeyLocal)("name"), strLocalDBFilePath) Then
                dbLocal.SQL = dicOriginTableSchema(varKeyLocal)("sql")
                dbLocal.DBPath = strLocalDBFilePath
                dbLocal.DoSQL_No_Transaction
            End If
        Next
        '次に、中身のコピーに進む（ここは重いはず）
        If chkBoxLastArea.Value Then
            '取得行数指定がある場合（こっちがデフォルト動作にしたい）
            isCollect = CopyDBTableRemote_To_Local(Table_JobDataPri & KishuNickName.KishuName, strLocalDBFilePath, longargLastNumberArea:=CLng(txtBoxAfterArea.Text))
        Else
            isCollect = CopyDBTableRemote_To_Local(Table_JobDataPri & KishuNickName.KishuName, strLocalDBFilePath)
        End If
    #End If
    vararrJobData = ReturnJobNumber_For_KanbanDivide(Table_JobDataPri & KishuNickName.KishuName, KishuNickName)
    cmbBoxJobNumber.ColumnCount = UBound(vararrJobData, 2) - LBound(vararrJobData, 2) + 1
    strListColumnWidth = GetColumnWidthString(vararrJobData, boolMaxLengthFind:=True)
    cmbBoxJobNumber.List = vararrJobData
    lblNowKishuNickName.Caption = cmbBoxKishuNickName.Text
End Sub
Private Sub cmbBoxJobNumber_Change()
    Dim vararrDivideChr As Variant
    Dim KishuLocal As typKishuInfo
    Dim strTableName As String
    Dim strJobNumber As String
    Dim strInputDate As String
    Dim intCounterRow As Integer
    Dim strDivideListColumnWidts As String
    On Error GoTo ErrorCatch
    'Job番号まで決まったら、指定のジョブ番号が2個有るかどうか確認する
    'Job番号ボックスには Job番号 InputDate 残り枚数の順で入ってるはず
    '↑リストから選んでもらえばいいか・・・
    '1列目Job番号、2列目InputInitialDateになってるはず
    '過去の分割情報を取得する
    '最初にタイトル行ありで帰ってくる
    '分割文字列 シート数（ダミー） 枚数 ラック数（ダミー） スタート履歴 エンド履歴 の順に帰ってくる
    If cmbBoxJobNumber.Text = "" Then
        Exit Sub
    End If
    '最初に過去結果リストボックスのお掃除
    listBoxExistingChr.Clear
    KishuLocal = GetKishuinfoByNickName(lblNowKishuNickName.Caption)
    strTableName = Table_JobDataPri & KishuLocal.KishuName
    strJobNumber = cmbBoxJobNumber.List(cmbBoxJobNumber.ListIndex, 0)
    strInputDate = cmbBoxJobNumber.List(cmbBoxJobNumber.ListIndex, 1)
    '右側の残りシート数/枚数ラベルを更新してやる
    lblRemainMaisuu.Caption = CStr(cmbBoxJobNumber.List(cmbBoxJobNumber.ListIndex, 2))
    lblRemainSheetQty.Caption = CStr(CLng(lblRemainMaisuu.Caption) / KishuLocal.MaiPerSheet)
    'ここまで来たら分割文字列以降をEnableにしてやる
    cmbBoxKanbanChr.Enabled = True
    txtBoxNewMaisuu.Enabled = True
    txtboxNewSheetQty.Enabled = True
    '次の看板文字列の候補をセットしてやる（Job無視）
    cmbBoxKanbanChr.Value = GetNextKanbanChrByTableName(strTableName)
    '次の開始履歴をセットしてやる
    lblNextRireki.Caption = GetNextKanbanRirekiByJobNumber(strTableName, strJobNumber, strInputDate)
    'フォーカス移動
    txtboxNewSheetQty.SetFocus
    vararrDivideChr = ReturnDivideChrByJobNumber(strTableName, strJobNumber, strInputDate)
    'なぜか幅が変わるので、手動設定
    listBoxExistingChr.Width = 524.24
    Sleep 2
    listBoxExistingChr.Width = 524.25
    'ここで過去の履歴なしの場合は、以後の処理を中止して過去結果リストボックスにそう表示してやる
    'データなしの場合は、新品のJobの可能性もあるので注意
'    If vararrDivideChr(0, 0) = "No Title" Then
    If UBound(vararrDivideChr, 2) < 2 Then
        '現時点でデータなし
        listBoxExistingChr.ColumnWidths = ""
        listBoxExistingChr.ColumnCount = 1
        listBoxExistingChr.AddItem "JOB分割履歴なし"
        Exit Sub
    End If
    'シート数とラック数はダミーの数が入ってるので、入れてやらないとダメ
    For intCounterRow = LBound(vararrDivideChr, 1) + 1 To UBound(vararrDivideChr, 1)
        vararrDivideChr(intCounterRow, 1) = CLng(vararrDivideChr(intCounterRow, 2) / KishuLocal.MaiPerSheet)
        vararrDivideChr(intCounterRow, 3) = CLng(Application.WorksheetFunction.RoundUp( _
                                            CSng(vararrDivideChr(intCounterRow, 1)) / CSng(KishuLocal.SheetPerRack), 0))
    Next intCounterRow
    strDivideListColumnWidts = GetColumnWidthString(vararrDivideChr, boolMaxLengthFind:=True)
    listBoxExistingChr.ColumnCount = UBound(vararrDivideChr, 2) - LBound(vararrDivideChr, 2) + 1
    listBoxExistingChr.ColumnWidths = strDivideListColumnWidts
    listBoxExistingChr.List = vararrDivideChr
    Exit Sub
ErrorCatch:
    Debug.Print "cmbBoxJobNumber_Change code: " & Err.Number & " Description: " & Err.Description
End Sub
Private Sub Clear_Exclude_KishuNickName(Optional ByVal boolExcludeJobNumber As Boolean)
    '機種通称名が選ばれた際に、他のものを初期化する
    If Not boolExcludeJobNumber Then
        cmbBoxJobNumber.Clear
    End If
    listBoxExistingChr.Clear
    listBoxExistingChr.ColumnCount = 1
'    lblNowKishuNickName.Caption = ""
    lblRemainMaisuu.Caption = ""
    lblRemainSheetQty.Caption = ""
    txtBoxNewMaisuu.Text = ""
    txtboxNewSheetQty.Text = ""
    listBoxExistingChr.Width = 524.5
End Sub
Private Sub txtBoxNewMaisuu_Change()
    Dim KishuLocal As typKishuInfo
    On Error GoTo ErrorCatch
    '自分のトコじゃなかったら無視する
    If Not ActiveControl.Name = txtBoxNewMaisuu.Name Then
        Exit Sub
    End If
    '空だったり、数字じゃなかったりしたら何もしない
    If txtBoxNewMaisuu.Text = "" Then
        txtboxNewSheetQty.Text = ""
        Exit Sub
    End If
    If Not IsNumeric(CLng(txtBoxNewMaisuu.Text)) Then
        Exit Sub
    End If
    If lblNowKishuNickName.Caption = "" Then
        Exit Sub
    End If
    KishuLocal = GetKishuinfoByNickName(lblNowKishuNickName.Caption)
    txtboxNewSheetQty.Text = CLng(txtBoxNewMaisuu.Text) / CLng(KishuLocal.MaiPerSheet)
    Exit Sub
ErrorCatch:
    Debug.Print "txtBoxNewMaisuu_Change code: " & Err.Number & " Description: " & Err.Description
    Exit Sub
End Sub
Private Sub txtBoxNewMaisuu_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    '枚数を入力した後、シート数を整数に切り上げする
    Dim KishuLocal As typKishuInfo
    KishuLocal = GetKishuinfoByNickName(lblNowKishuNickName.Caption)
    txtboxNewSheetQty.Text = Application.WorksheetFunction.RoundUp(txtboxNewSheetQty.Text, 0)
    txtBoxNewMaisuu.Text = CLng(txtboxNewSheetQty.Text) * CLng(KishuLocal.MaiPerSheet)
    If CLng(txtBoxNewMaisuu.Text) > CLng(lblRemainMaisuu.Caption) Then
        '計算してみた結果残り枚数を超えるようなら、残り枚数をシートで割って、小数切り捨てにしたのをシート数に入れてやる
        txtboxNewSheetQty.Text = Int(CLng(lblRemainMaisuu.Caption) / CLng(KishuLocal.MaiPerSheet))
        txtBoxNewMaisuu.Text = CLng(txtboxNewSheetQty.Text) * CLng(KishuLocal.MaiPerSheet)
    End If
End Sub
Private Sub txtboxNewSheetQty_Change()
    Dim KishuLocal As typKishuInfo
    On Error GoTo ErrorCatch
    '自分のとこじゃない場所がアクティブになってたら何もしない
    If Not ActiveControl.Name = txtboxNewSheetQty.Name Then
        Exit Sub
    End If
    '空だったり、数字じゃなかったりしたら何もしない
    If txtboxNewSheetQty.Text = "" Then
        txtBoxNewMaisuu.Text = ""
        Exit Sub
    End If
    If Not IsNumeric(CLng(txtboxNewSheetQty.Text)) Then
        Exit Sub
    End If
    If lblNowKishuNickName.Caption = "" Then
        Exit Sub
    End If
    KishuLocal = GetKishuinfoByNickName(lblNowKishuNickName.Caption)
    txtBoxNewMaisuu.Text = CLng(txtboxNewSheetQty.Text) * CLng(KishuLocal.MaiPerSheet)
    Exit Sub
ErrorCatch:
    Debug.Print "txtBoxNewSheet_Change code: " & Err.Number & " Description: " & Err.Description
    Exit Sub
End Sub
Private Sub txtboxNewSheetQty_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    '最大数を超えてないかチェックする
    Dim KishuLocal As typKishuInfo
    On Error GoTo ErrorCatch
    KishuLocal = GetKishuinfoByNickName(lblNowKishuNickName.Caption)
    If CLng(txtboxNewSheetQty.Text) > CLng(lblRemainSheetQty.Caption) Then
        txtboxNewSheetQty.Text = lblRemainSheetQty.Caption
        txtBoxNewMaisuu.Text = CLng(txtboxNewSheetQty.Text) * CLng(KishuLocal.MaiPerSheet)
    End If
    Exit Sub
ErrorCatch:
    Debug.Print "txtboxNewSheetQty Exit code: "; Err.Number & " Description: " & Err.Description
    Exit Sub
End Sub
Private Sub UserForm_Initialize()
    'ローカルコピーモードの時は、テンポラリディレクトリを作成する
    #If READLOCAL = 1 Then
        Dim fsoLocal As FileSystemObject
        Set fsoLocal = New FileSystemObject
        strLocalPath = ThisWorkbook.Path & "\" & LocalTempDBDir
        strLocalDBFilePath = strLocalPath & "\" & LocalDBName
        If Not fsoLocal.FolderExists(strLocalPath) Then
            'ディレクトリ無ければ作成する
            MkDir strLocalPath
        End If
        Set fsoLocal = Nothing
    #End If
    '看板分割フォーム初期化
    Dim dbKanban As clsSQLiteHandle
    Dim varArrKishuNickName As Variant
    Dim intCounterKishu As Integer
    Dim byteChrCodeCounter As Byte
    '機種（通称名）一覧を取得する
    Set dbKanban = New clsSQLiteHandle
    dbKanban.SQL = "SELECT " & Kishu_KishuNickname & " FROM " & Table_Kishu
    dbKanban.DoSQL_No_Transaction
    varArrKishuNickName = dbKanban.RS_Array(boolPlusTytle:=False)
    Set dbKanban = Nothing
    '機種名コンボボックスに追加してやる
    cmbBoxKishuNickName.List = varArrKishuNickName
    '看板分割文字列ボックスにA-Zを追加
    For byteChrCodeCounter = MIN_Kanban_ChrCode To MAX_Kanban_ChrCode
        cmbBoxKanbanChr.AddItem Chr(byteChrCodeCounter)
    Next byteChrCodeCounter
    btnQRRead.SetFocus
End Sub
Private Sub UserForm_Terminate()
    #If READLOCAL = 1 Then
        'todo
        'ローカルコピー使用時は一時ファイルを・・・削除する？しない？
        '接続しっぱなしになってるLocalDBをクローズ
        Set dbLocal = Nothing
    #End If
End Sub