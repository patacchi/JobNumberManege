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


Private Sub btnInputRirekiNumber_Click()
    'ジョブ番号・履歴の登録処理
    '以下はDBから引っ張ってくる
    Dim dicKishuInfo As New Scripting.Dictionary
    Dim byteRenbannKeta As Byte         '連番部分の桁数（Right関数の引数）
    Dim byteRirekiKetaCount As Byte     '履歴の桁数トータル
    
    Dim strHeader As String             'ヘッダ
    Dim longRenbann As Long             '履歴連番部分
    Dim strRireki As String             'ヘッダ+連番
    Dim longInputRow As Long            'シート入力時の入力行数
    Dim longLocalCounter As Long        'ループ処理用カウンター
    Dim nameRirekiKetasuu As Name       '履歴桁数の名前定義
    Dim arryKishuHeader() As Variant    'ジョブ情報_機種テーブルから、機種名一覧を受け取る
    
    If txtboxJobNumber.Text = Empty Or _
        txtboxMaisuu.Text = Empty Or _
        txtboxStartRireki.Text = Empty Then
        MsgBox ("空白の項目があります。確認してください")
        Exit Sub
    End If
    
    If Len(txtboxStartRireki.Text) > constMaxRirekiKetasuu Then
        MsgBox ("履歴の桁数が" & constMaxRirekiKetasuu & "桁を超えています。処理を中止します。")
        Exit Sub
    End If

    If txtboxMaisuu.Text < 1 Then
        MsgBox ("枚数には1以上の整数を入力して下さい")
        Exit Sub
    End If

    
    'ヘッダーテーブルより履歴入力部分から機種名（履歴構成）を引っ張ってくる
    '先頭の文字列一致でいいかな？
'    'カレントディレクトリ変更（DBディレクトリへ）
'    If Not ChcurrentforDB() Then
'        MsgBox ("DBディレクトリ認識失敗。処理を中断します。")
'        Exit Sub
'    End If
'
'    'DBファイルの存在有無の確認（なければ作る）
'    If Not IsDBFileExist() Then
'        MsgBox ("失敗・・・・")
'    End If
    '機種名一覧を受け取る処理
    
    '返ってくるリストは、機種ヘッダ、機種名、トータル桁数、連番桁数の順番
    'arryKishuHeader = KishuList()
    If Not arryKishuHeader Then
        '機種情報見つからなかったので、機種登録から
        
    End If
    
    'フォームに入力された文字の先頭と履歴ヘッダ（機種判別用）が一致するか調べる
    '見つからない場合は機種登録画面を(todo)
    For longLocalCounter = 0 To UBound(arryKishuHeader)
        
    Next longLocalCounter
    
End Sub

