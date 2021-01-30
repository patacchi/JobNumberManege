VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJobNumberInput 
   Caption         =   "ジョブ番号・履歴登録画面"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7050
   OleObjectBlob   =   "frmJobNumberInput.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmJobNumberInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Option Base 1

Private Sub frmJobNumberInput_Initialize()
    'フォーム初期化（全部消すだけ）
    txtboxJobNumber.Text = ""
    txtboxMaisuu.Text = ""
    txtboxStartRireki = ""
    labelZuban.Caption = ""
    btnQRFormShow.SetFocus
End Sub


Private Sub btnInputRirekiNumber_Click()
    'ジョブ番号・履歴の登録処理
    Dim KishuInfoLocal As typKishuInfo
    Dim isCollect As Boolean
    Dim sqlbJobInput As clsSQLStringBuilder
    On Error GoTo ErrorCatch

    If txtboxJobNumber.Text = Empty Or _
        txtboxMaisuu.Text = Empty Or _
        txtboxStartRireki.Text = Empty Then
        MsgBox ("空白の項目があります。確認してください")
        Exit Sub
    End If

    If CLng(txtboxMaisuu.Text) < 1 Then
        MsgBox ("枚数には1以上の整数を入力して下さい")
        Exit Sub
    End If
    
    'スタート履歴からKishuInfoを引っ張ってくる
    KishuInfoLocal = getKishuInfoByRireki(txtboxStartRireki.Text)
    Set sqlbJobInput = New clsSQLStringBuilder
    With sqlbJobInput
        .JobNumber = CStr(txtboxJobNumber.Text)
        .FieldArray = arrFieldList_JobData
        .StartRireki = CStr(txtboxStartRireki.Text)
        .Maisu = CLng(txtboxMaisuu.Text)
        .RenbanKeta = KishuInfoLocal.RenbanKetasuu
        .TableName = Table_JobDataPri & KishuInfoLocal.KishuName
    End With
    Set sqlbJobInput.FieldType = GetFieldTypeNameByTableName(sqlbJobInput.TableName)
    If Not Len(txtboxStartRireki.Text) = KishuInfoLocal.TotalRirekiketa Then
        MsgBox "履歴の桁数が登録されている機種名：" & KishuInfoLocal.KishuName & " の " & _
                KishuInfoLocal.TotalRirekiketa & " 桁と違います。処理を中止します。"
                GoTo CloseAndExit
    End If
    
    isCollect = sqlbJobInput.CreateInsertSQL(boolCheckLastRireki:=True)
    If Not isCollect Then
        MsgBox "ジョブ登録中に何かあったようです"
        GoTo CloseAndExit
        Exit Sub
    End If
    
    MsgBox "ジョブ登録完了"
    frmJobNumberInput_Initialize
    GoTo CloseAndExit
    Exit Sub
CloseAndExit:
    Set sqlbJobInput = Nothing
    Exit Sub
ErrorCatch:
    Debug.Print "ImputRireki Erro code: " & Err.Number & "Description: " & Err.Description
    Set sqlbJobInput = Nothing
    Exit Sub
End Sub

Private Sub btnQRFormShow_Click()
    'QRコード読み取りフォーム表示
'    frmJobNumberInput.Hide
    frmQRAnalyze.Show
End Sub
