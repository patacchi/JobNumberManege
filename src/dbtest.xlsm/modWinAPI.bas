Attribute VB_Name = "modWinAPI"
Option Explicit
'リサイズ実装のためWin32API使用
'const
Public Const GWL_STYLE As Long = (-16)                     'ウィンドウスタイルのハンドラ番号
Public Const WS_MAXIMIZEBOX As Long = &H10000  'ウィンドウスタイルで最大化ボタンをつける
Public Const WS_MINIMIZEBOX As Long = &H20000  'ウィンドウスタイルで最小化ボタンを付ける
Public Const WS_THICKFRAME As Long = &H40000   'ウィンドウスタイルでサイズ変更をつける
Public Const WS_SYSMENU As Long = &H80000      'ウィンドウスタイルでコントロールメニューボックスをもつウィンドウを作成する
'-----Windows API宣言-----
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#If Win64 Then
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
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