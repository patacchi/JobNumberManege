VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQLTest 
   Caption         =   "SQLテスト"
   ClientHeight    =   8625.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13785
   OleObjectBlob   =   "frmSQLTest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSQLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'リサイズ実装のためWin32API使用
'const
Option Explicit
Private Const GWL_STYLE As Long = (-16)                     'ウィンドウスタイルのハンドラ番号
Private Const WS_MAXIMIZEBOX As Long = &H10000  'ウィンドウスタイルで最大化ボタンをつける
Private Const WS_MINIMIZEBOX As Long = &H20000  'ウィンドウスタイルで最小化ボタンを付ける
Private Const WS_THICKFRAME As Long = &H40000   'ウィンドウスタイルでサイズ変更をつける
Private Const WS_SYSMENU As Long = &H80000      'ウィンドウスタイルでコントロールメニューボックスをもつウィンドウを作成する
'-----Windows API宣言-----
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'フォームに最大化・リサイズ機能を追加する。
Public Sub FormResize()
        Dim hwnd As LongPtr
        Dim WndStyle As LongPtr

    'ウィンドウハンドルの取得
    hwnd = GetActiveWindow()
    'ウィンドウのスタイルを取得
    WndStyle = GetWindowLongPtr(hwnd, GWL_STYLE)
    '最大・最小・サイズ変更を追加する
    WndStyle = WndStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_SYSMENU

    Call SetWindowLongPtr(hwnd, GWL_STYLE, WndStyle)
End Sub


Private Sub UserForm_Activate()
    'リサイズ機能追加
    Call FormResize
End Sub
Private Sub UserForm_Resize()
    'フォームリサイズ時に、中のリストボックスもサイズ変更してやる
    Dim intListHeight As Integer
    Dim intListWidth As Integer
    
    intListHeight = Me.InsideHeight - listBoxSQLResult.Top * 2
    intListWidth = Me.InsideWidth - (txtboxSQLText.Left * 2) - txtboxSQLText.Width - (listBoxSQLResult.Left - txtboxSQLText.Width - txtboxSQLText.Left)
    If (intListHeight > 0 And intListWidth > 0) Then
        listBoxSQLResult.Height = intListHeight
        listBoxSQLResult.Width = intListWidth
    End If

End Sub
Private Sub btnBulkDataInput_Click()
    Dim strSQL
    Randomize
    frmBulkInsertTest.Show
    'ある範囲の乱数の発生のさせ方
    'Int((範囲上限値 - 範囲下限値 + 1) * Rnd + 範囲下限値)

End Sub
Private Sub btnSQLGo_Click()
    'エラーチェックとかほとんどなし
    'テキストボックスに入れたSQLを実行するフォームっぽいの
    If txtboxSQLText.Text = "" Then
        MsgBox "空白はちょっと・・・"
        Exit Sub
    End If
    Dim dbSQLite3 As clsSQLiteHandle
    Dim varRetValue As Variant
    Dim isCollect As Boolean
    Set dbSQLite3 = New clsSQLiteHandle
    IsDBFileExist
    isCollect = dbSQLite3.DoSQL_No_Transaction(txtboxSQLText.Text)
    If isCollect Then
        varRetValue = dbSQLite3.RS_Array(boolPlusTytle:=True)
    Else
        'エラーがあった場合の処理・・・なんだけど
        'エラーメッセージをそのまま表示すればいいのでは・・・
        varRetValue = dbSQLite3.RS_Array(boolPlusTytle:=True)
    End If
    If VarType(varRetValue) = vbEmpty Then
        listBoxSQLResult.Clear
        listBoxSQLResult.AddItem "データなし"
        Exit Sub
    End If
    
    Set dbSQLite3 = Nothing
    With listBoxSQLResult
        .ColumnCount = UBound(varRetValue, 2) - LBound(varRetValue, 2) + 1
        '.ColumnWidths = "50;50;50;50;50;50;50;50"
        '.List = Join(varRetValue)
        .List = varRetValue
        '.AddItem (varRetValue(1)(1))
    End With
End Sub
Private Sub listBoxSQLResult_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'リストダブルクリックしたらクリップボードにコピーしてみおよう
    Dim objDataObj As DataObject
    Dim intCounterColumn As Integer
    Dim strListText As String
    Set objDataObj = New DataObject
        objDataObj.SetText (listBoxSQLResult.List(listBoxSQLResult.ListIndex))
        objDataObj.PutInClipboard
        strListText = ""
        For intCounterColumn = 0 To listBoxSQLResult.ColumnCount - 1
            If IsNull(listBoxSQLResult.List(listBoxSQLResult.ListIndex, intCounterColumn)) Then
                'Nullの場合はNULLって入れてやろう
                strListText = strListText & " NULL"
            Else
                strListText = strListText & " " & CStr(listBoxSQLResult.List(listBoxSQLResult.ListIndex, intCounterColumn))
            End If
        Next intCounterColumn
        LTrim (strListText)
        MsgBox strListText
        Debug.Print strListText
End Sub
