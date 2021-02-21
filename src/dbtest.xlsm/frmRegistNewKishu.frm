VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistNewKishu 
   Caption         =   "新機種登録"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12075
   OleObjectBlob   =   "frmRegistNewKishu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmRegistNewKishu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnAlreadyDataSet_Click()
    'リストボックスで選択されている機種情報を左のあちこちに張り付ける
'    Dim intKishuCount As Integer
    Dim KishuInfoLocal As typKishuInfo
    If ListBoxAlreadyKishu.ListIndex = -1 Then
        MsgBox "左のリストのうちどれかを選んでからもう一度クリックして下さい。"
        Exit Sub
    End If
'    If (Not arrKishuInfoGlobal) = -1 Then
'        'ぐろーばるkishuinfoが行方不明なので設定
'        Call GetAllKishuInfo_Array
'    End If
'    If boolNoTableKishuRecord Then
'        Exit Sub
'    End If
    '選択された機種名（List,1）よりKishuInfoを引っ張る
    KishuInfoLocal = GetKishuinfoByKishuName(ListBoxAlreadyKishu.List(ListBoxAlreadyKishu.ListIndex, 1))
    'Uniqueフィールドは空白の場合のみ入力、空フィールドチェックで上書きされる危険性があるため
    'UpdateModeになってる時は、フォーム初期化時に値をセットしてるのでいじらないようにする
    'しかし本当の新機種登録の時に、間違ってテンプレ選んだ場合もあるので、orで複数条件設定
    If labelKishuHeader.Caption = "" Or chkBoxUpdateMode.Value = False Then
        txtboxKishuHeader.Text = Len(KishuInfoLocal.KishuHeader)
    End If
    '機種名は、QRコードの情報を引き継いでいる可能性があるので、空白の場合のみ入力
    If txtboxKishuName.Text = "" Then
        txtboxKishuName.Text = KishuInfoLocal.KishuName
    End If
    If txtBoxKishuNickName.Text = "" Or chkBoxUpdateMode.Value = False Then
        txtBoxKishuNickName.Text = KishuInfoLocal.KishuNickName
    End If
    '履歴桁数のトータルは動かしちゃダメ
'   txtboxTotalRirekiKetasuu.Text = arrKishuInfoGlobal(intKishuCount).TotalRirekiketa
    If txtboxRenbanketasuu.Text = "" Or chkBoxUpdateMode.Value = False Then
        txtboxRenbanketasuu.Text = KishuInfoLocal.RenbanKetasuu
    End If
    If txtBoxMaiPerSheet.Text = "" Or chkBoxUpdateMode.Value = False Then
        txtBoxMaiPerSheet.Text = KishuInfoLocal.MaiPerSheet
    End If
    If txtBoxSheetPerRack.Text = "" Or chkBoxUpdateMode.Value = False Then
        txtBoxSheetPerRack.Text = KishuInfoLocal.SheetPerRack
    End If
    If txtBoxBarcodeReadNumber.Text = "" Or chkBoxUpdateMode.Value = False Then
        txtBoxBarcodeReadNumber.Text = KishuInfoLocal.BarcordReadNumber
    End If
    txtboxKishuHeader.SetFocus
    Exit Sub
End Sub
Private Sub btnCancel_Click()
    'とりあえず処理を中止する
    MsgBox "キャンセルボタンが押されたため、処理を中止します。"
    boolRegistOK = False
    Unload Me
End Sub
Private Sub btnregistNewKishu_Click()
    Dim longRecordCount As Long
    Dim longMsgBoxReturn As Long
    On Error GoTo ErrorCatch
    If labelKishuHeader.Caption = "" Or txtboxKishuName.Text = "" Or txtboxTotalRirekiKetasuu.Text = "" Or txtboxRenbanketasuu.Text = "" Or txtBoxKishuNickName.Text = "" _
        Or txtBoxMaiPerSheet.Text = "" Or txtBoxSheetPerRack.Text = "" Or txtBoxBarcodeReadNumber.Text = "" Then
        MsgBox "未入力箇所があります。入力し直してください"
        Exit Sub
    End If
    '既存の物と重複してないかチェックしてみる（簡易版）
    Select Case chkBoxUpdateMode.Value
    Case False
        '通常の新規登録の場合
        longRecordCount = GetRecordCountSimple(Table_Kishu, Kishu_Header, "LIKE """ & labelKishuHeader.Caption & """")
        If longRecordCount >= 1 Then
            MsgBox "機種ヘッダの重複があるようです。入力内容を確認して下さい。"
            txtboxKishuHeader.SetFocus
            Exit Sub
        End If
        longRecordCount = GetRecordCountSimple(Table_Kishu, Kishu_KishuName, "LIKE """ & txtboxKishuName.Text & """")
        If longRecordCount >= 1 Then
            MsgBox "機種名で重複があるようです。入力内容を確認して下さい。"
            txtboxKishuName.SetFocus
            Exit Sub
        End If
        longRecordCount = GetRecordCountSimple(Table_Kishu, Kishu_KishuNickname, "LIKE """ & txtBoxKishuNickName.Text & """")
        If longRecordCount >= 1 Then
            MsgBox "機種通称名で重複があるようです。入力内容を確認して下さい。"
            txtBoxKishuNickName.SetFocus
            Exit Sub
        End If
    Case True
        'Update Modeの時
        '後で何かやりたくなったらここで
    End Select
    On Error Resume Next
    If Not IsNumeric(CLng(labelRenban.Caption)) Then
        Debug.Print "InNumeric RenbanCaption code: " & Err.Number & " Descriptoin: " & Err.Description
        If Err.Number = 13 Then
            '13=型が一致しません
            MsgBox "連番部分に数値以外が混入しているようです。連番の桁数を確認して下さい。"
            Exit Sub
        ElseIf Err.Number = 6 Then
            '6 = オーバーフローしました
            MsgBox "32bitExcelで扱える数字の桁数を超えています。連番の桁数を確認して下さい。"
            Exit Sub
        End If
        txtboxRenbanketasuu.SetFocus
        Exit Sub
    End If
    If Err.Number <> 0 Then
        MsgBox "連番部分に数値以外が混入しているようです。連番の桁数を確認して下さい。"
        txtboxRenbanketasuu.SetFocus
        Exit Sub
    End If
    '連番、機種ヘッダ桁数がトータル桁数超えてないかどうか
    If CInt((txtboxKishuHeader.Text) + CInt(txtboxRenbanketasuu.Text)) > CInt(txtboxTotalRirekiKetasuu.Text) Then
        MsgBox "履歴ヘッダの桁数と連番桁数の合計が履歴のトータル桁数を超えています。"
        txtboxKishuHeader.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrorCatch
    If CByte(txtboxTotalRirekiKetasuu.Text) > constMaxRirekiKetasuu Then
        longMsgBoxReturn = MsgBox(prompt:="履歴の桁数が " & constMaxRirekiKetasuu & "桁を超えていますが、続行しますか？", Buttons:=vbYesNo)
        If longMsgBoxReturn = vbNo Then
            boolRegistOK = False
            Unload Me
            Exit Sub
        End If
    End If
    Select Case chkBoxUpdateMode.Value
    Case True
        boolRegistOK = registNewKishu_to_KishuTable(labelKishuHeader.Caption, txtboxKishuName.Text, txtBoxKishuNickName.Text, _
                        CByte(txtboxTotalRirekiKetasuu.Text), CByte(txtboxRenbanketasuu.Text), _
                        CByte(txtBoxMaiPerSheet.Text), CByte(txtBoxSheetPerRack.Text), CByte(txtBoxBarcodeReadNumber.Text), boolUpdateMode:=True)
        'Update Mode解除
        chkBoxUpdateMode.Value = False
    Case False
        boolRegistOK = registNewKishu_to_KishuTable(labelKishuHeader.Caption, txtboxKishuName.Text, txtBoxKishuNickName.Text, _
                        CByte(txtboxTotalRirekiKetasuu.Text), CByte(txtboxRenbanketasuu.Text), _
                        CByte(txtBoxMaiPerSheet.Text), CByte(txtBoxSheetPerRack.Text), CByte(txtBoxBarcodeReadNumber.Text))
    End Select
    If boolRegistOK Then
        'noKishuフラグが立ってたらひっこめる
        boolNoTableKishuRecord = False
        'グローバルのKishuInfoを更新してやる
        Call GetAllKishuInfo_Array
        Unload Me
        Exit Sub
    Else
        MsgBox "機種登録作業でエラーが発生しました"
        Debug.Print "機種登録フラグNGにより終了"
        Unload Me
        Exit Sub
    End If
    Exit Sub
ErrorCatch:
    MsgBox "機種登録中にエラーが発生したようです。処理を中止します"
    Debug.Print "btnRegistNewKishu_click code: " & Err.Number & " Description: " & Err.Description
    boolRegistOK = False
    Unload Me
    Exit Sub
End Sub
Private Sub txtboxKishuHeader_Change()
    '機種ヘッダの桁数が変化したら、横のラベルに履歴の左端からの指定文字数を入れてやる
    Dim intStringCount As Integer
    If txtboxKishuHeader.Text = "" Then
        intStringCount = 0
    Else
        intStringCount = CInt(txtboxKishuHeader.Text)
    End If
    If intStringCount >= Len(strRegistRireki) Then
        intStringCount = Len(strRegistRireki)
    End If
    labelKishuHeader.Caption = Left(strRegistRireki, intStringCount)
End Sub
Private Sub txtboxRenbanketasuu_Change()
    Dim intStringCount As Integer
    If txtboxRenbanketasuu.Text = "" Then
        intStringCount = 0
    Else
        intStringCount = CInt(txtboxRenbanketasuu.Text)
    End If
    If intStringCount >= Len(strRegistRireki) Then
        intStringCount = Len(strRegistRireki)
    End If
    labelRenban.Caption = Right(strRegistRireki, intStringCount)
End Sub
Private Sub txtboxKishuHeader_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtboxRenbanketasuu_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub
Private Sub UserForm_Initialize()
    '初期化処理として・・・
    Dim strRireki As String
    Dim byteLocalCounter As Byte
    Dim intCountKishuInfo As Integer
    Dim varArrRegestedKishu As Variant
    Dim strUpperRubi As String  '履歴上部表示用ルビ
    Dim strMiddle As String     '履歴真ん中
    Dim strLowerRubi As String  '履歴下部ルビ
    Dim strHantei As String '履歴判定用テキスト、上下に数字のルビを
'    Dim dbKishuList As clsSQLiteHandle
    strRireki = strRegistRireki
    strMiddle = ""
    strHantei = ""
    strUpperRubi = ""
    strLowerRubi = ""
    'まずは左が1で右が履歴桁数になるルビを、ついでに真ん中も
    For byteLocalCounter = 1 To Len(strRireki)
        strUpperRubi = strUpperRubi & byteLocalCounter
        strUpperRubi = strUpperRubi & Space(3 - Len(CStr(byteLocalCounter)))
        strMiddle = strMiddle & Mid(strRireki, byteLocalCounter, 1)
        strMiddle = strMiddle & Space(2)
    Next byteLocalCounter
    strUpperRubi = RTrim(strUpperRubi)
    strMiddle = RTrim(strMiddle)
    byteLocalCounter = Len(strRireki)
    '次に下部のルビを
    Do While byteLocalCounter >= 1
        strLowerRubi = strLowerRubi & byteLocalCounter
        strLowerRubi = strLowerRubi & Space(3 - Len(CStr(byteLocalCounter)))
        byteLocalCounter = byteLocalCounter - 1
    Loop
    strLowerRubi = RTrim(strLowerRubi)
    '判定用テキスト合体
    strHantei = strUpperRubi & vbCrLf & strMiddle & vbCrLf & strLowerRubi
    txtboxHanteiRireki = strHantei
    txtboxTotalRirekiKetasuu.Text = Len(strRireki)
    '既存機種リストボックスの初期化
    If Not IsTableExist(Table_Kishu) Then
        MsgBox "機種テーブルがありません。新規作成します"
        InitialDBCreate
    End If
    '機種テーブルより、KishuName、KishuNicknameをとってきて表示してやる
    If (Not arrKishuInfoGlobal) = -1 Then
        'グローバルKishuInfoが初期化されてないので、やってみる
        Call GetAllKishuInfo_Array
    End If
'    dbKishuList.SQL = "SELECT " & Kishu_Header & " as 機種ヘッダ , " & _
'                        Kishu_KishuName & " as 機種名 , " & _
'                        Kishu_KishuNickname & " as 通称名 FROM " & Table_Kishu
'    Call dbKishuList.DoSQL_No_Transaction
    ReDim varArrRegestedKishu(UBound(arrKishuInfoGlobal) + 1, 2)
    'タイトル
    varArrRegestedKishu(0, 0) = "機種ヘッダ"
    varArrRegestedKishu(0, 1) = "機種名"
    varArrRegestedKishu(0, 2) = "機種通称名"
    'KishuInfoGlobalより情報を入れてやる
    For intCountKishuInfo = LBound(arrKishuInfoGlobal) + 1 To UBound(arrKishuInfoGlobal) + 1
        varArrRegestedKishu(intCountKishuInfo, 0) = arrKishuInfoGlobal(intCountKishuInfo - 1).KishuHeader
        varArrRegestedKishu(intCountKishuInfo, 1) = arrKishuInfoGlobal(intCountKishuInfo - 1).KishuName
        varArrRegestedKishu(intCountKishuInfo, 2) = arrKishuInfoGlobal(intCountKishuInfo - 1).KishuNickName
    Next intCountKishuInfo
    ListBoxAlreadyKishu.ColumnCount = 3
    ListBoxAlreadyKishu.List = varArrRegestedKishu
    '履歴のトータル桁数を設定し、そこを編集不可に
    If strRegistRireki = "" Then
        txtboxTotalRirekiKetasuu.Text = 0
    Else
        txtboxTotalRirekiKetasuu.Text = CStr(Len(strRegistRireki))
    End If
    txtboxTotalRirekiKetasuu.Enabled = False
    'QRコードから読み取った図番がある場合は、機種名に適用
    If Not QRField.Zuban = "" Then
        txtboxKishuName.Text = QRField.Zuban
    End If
    '多分機種登録されてないからここに来たんだろうと言う事で
    boolRegistOK = False
    MsgBox "機種が登録されていない、又は機種登録時未入力箇所があったようなので登録画面に移行します。"
End Sub
Private Sub UserForm_Terminate()
    '終了処理
'    strQRZuban = ""
    strRegistRireki = ""
End Sub