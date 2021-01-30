VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQRAnalyze 
   Caption         =   "QRコード読み取り"
   ClientHeight    =   2280
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   4305
   OleObjectBlob   =   "frmQRAnalyze.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmQRAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    txtboxQRString.Text = ""
    frmQRAnalyze.Hide
    'formJobNumberInput.Show
    
End Sub
Private Sub txtboxQRString_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    '何か入力されたら反応して
    Dim strSplit As Variant
    Dim strJobNumber As String
    Dim intMaisuu As Integer
    Dim intCount As Integer
    Dim strBuf As String
    If txtboxQRString.Text = "" Then
        Debug.Print "String Empty"
        Exit Sub
    End If
    On Error GoTo ErrorCatcch
    strSplit = Split(txtboxQRString.Text, ",")
    If UBound(strSplit) < 4 Then
        '要素数が4以下の場合は指示書のQRコードじゃないっぽい
        MsgBox "指示書のQRコード以外が読み込まれた可能性があります。"
        txtboxQRString.Text = ""
        frmQRAnalyze.Hide
        frmQRAnalyze.Show
        Exit Sub
    End If
    intMaisuu = CInt(strSplit(3))
    'ジョブ番号の空白の連続をマージする
    strBuf = ""
    For intCount = 1 To Len(strSplit(0))
        Select Case intCount
        Case 1
            '1文字目は素直にバッファに入れてやる
            strBuf = strBuf & Mid(strSplit(0), intCount, 1)
        Case Else
            If Not Mid(strSplit(0), intCount, 1) = " " Then
                '空白以外だったら素直にバッファに入れてやる
                strBuf = strBuf & Mid(strSplit(0), intCount, 1)
            Else
                '空白だった場合は、直前のバッファの終端文字が空白かどうかで入れるか判断
                If Right(strBuf, 1) = " " Then
                    'ラストが空白だったら今回のループでは何もしない
                Else
                    'ラストが空白じゃんない時は1回目のスペースとしてバッファに入れる
                    strBuf = strBuf & Mid(strSplit(0), intCount, 1)
                End If
            End If
            
        End Select
        strJobNumber = strBuf
    Next intCount
    formJobNumberInput.textBoxJobNumber.Text = strJobNumber
    formJobNumberInput.textboxMisuu = intMaisuu
    formJobNumberInput.Show
    formJobNumberInput.textboxStartRireki.SetFocus
ErrorCatcch:
    '基本的にエラー発生したら何もしないや
    Debug.Print "Errror code: " & Err.Number & "Description: " & Err.Description
    txtboxQRString.Text = ""
    txtboxQRString.Enabled = False
    frmQRAnalyze.Hide
End Sub

Private Sub UserForm_Activate()
    txtboxQRString.Enabled = True
    txtboxQRString.SetFocus
End Sub

