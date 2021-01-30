VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistNewKishu 
   Caption         =   "新機種登録（日本語入力不可）"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8805.001
   OleObjectBlob   =   "frmRegistNewKishu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmRegistNewKishu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnRegistNewKishu_Click()
    Dim boolRegistOK As Boolean
    If txtboxKishuHeader.Text = "" Or txtboxKishuName.Text = "" Or txtboxTotalRirekiKetasuu.Text = "" Or txtboxRenbanketasuu.Text = "" Or txtBoxKishuNickName.Text = "" Then
        MsgBox "未入力箇所があります。入力し直してください"
        Exit Sub
    End If
    
    boolRegistOK = registNewKishu(txtboxKishuHeader.Text, txtboxKishuName.Text, txtBoxKishuNickName.Text, _
                    CByte(txtboxTotalRirekiKetasuu.Text), CByte(txtboxRenbanketasuu.Text))
    
End Sub

Private Sub txtboxRenbanketasuu_Change()
    Dim intStringCount As Integer
    intStringCount = CInt(txtboxRenbanketasuu.Text)
    If intStringCount >= Len(frmJobNumberInput.txtboxStartRireki.Text) Then
        intStringCount = Len(frmJobNumberInput.txtboxStartRireki.Text)
    End If
    labelRenban.Caption = Right(frmJobNumberInput.txtboxStartRireki.Text, intStringCount)

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
    Dim strUpperRubi As String  '履歴上部表示用ルビ
    Dim strMiddle As String     '履歴真ん中
    Dim strLowerRubi As String  '履歴下部ルビ
    Dim strHantei As String '履歴判定用テキスト、上下に数字のルビを
    strRireki = frmJobNumberInput.txtboxStartRireki.Text
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

End Sub
