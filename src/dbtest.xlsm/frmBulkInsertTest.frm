VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBulkInsertTest 
   Caption         =   "バルクインサートテスト"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600.001
   OleObjectBlob   =   "frmBulkInsertTest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmBulkInsertTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    'テスト用にデータを何個か入れてやる
    Dim KishuInfo As typKishuInfo
    txtboxStartRireki.Text = "Test0015T00000000001"
    txtboxJobNumber.Text = "TT00121"
    txtboxMaisuu.Text = 10000
    txtBoxFieldList.Text = Job_Number & "," & Job_RirekiHeader & "," & Job_RirekiNumber & "," & Job_Rireki
    KishuInfo = getKishuInfoByRireki(txtboxStartRireki.Text)
    If chkBoxInputNextRireki Then
        '最新履歴入力にチェックが入ってた場合
        '履歴とJob番号を自動入力する
        txtboxStartRireki.Text = GetNextRireki(Table_JobDataPri & KishuInfo.KishuName)
        txtboxJobNumber.Text = GetNextJobNumber_ForBulk(Table_JobDataPri & KishuInfo.KishuName)
    End If
    txtboxTableName.Text = Table_JobDataPri & KishuInfo.KishuName
End Sub
Private Sub btnGoInsert_Click()
    'Insertテスト
    Dim isCollect  As Boolean
    Dim strLastRireki As String
    Dim vararrField As Variant
    Dim dbSQLite3 As clsSQLiteHandle
    Set dbSQLite3 = New clsSQLiteHandle
    Dim KishuInfo As typKishuInfo
    Dim sqlbBulkSQL As clsSQLStringBuilder
    On Error GoTo ErrorCatch
    Set sqlbBulkSQL = New clsSQLStringBuilder
    KishuInfo = getKishuInfoByRireki(txtboxStartRireki.Text)
    '拾ってきた機種情報を元にいろいろごにょごにょ
    txtboxTableName.Text = Table_JobDataPri & KishuInfo.KishuName
    With sqlbBulkSQL
        .startRireki = txtboxStartRireki.Text
        .JobNumber = txtboxJobNumber.Text
        .Maisu = CLng(txtboxMaisuu.Text)
        .TableName = txtboxTableName.Text
        '.FieldArray = Split(txtBoxFieldList.Text, ",")
        .FieldArray = arrFieldList_JobData
        .RenbanKeta = KishuInfo.RenbanKetasuu
    End With
    Set sqlbBulkSQL.FieldType = GetFieldTypeNameByTableName(txtboxTableName.Text)
    Set dbSQLite3 = Nothing
    isCollect = sqlbBulkSQL.CreateInsertSQL()
    If Not isCollect Then
        MsgBox "バルクインサートテスト最後に何かあったっぽい？"
        GoTo ErrorCatch
    End If
    UserForm_Initialize
    Exit Sub
ErrorCatch:
    Debug.Print "btnGOInsert_Click code: " & Err.Number & "Description: " & Err.Description
End Sub