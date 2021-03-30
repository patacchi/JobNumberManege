Attribute VB_Name = "modWinAPI"
Option Explicit
'UNC対応のため、Win32API使用
Public Declare PtrSafe Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As LongPtr
'日付をミリ秒単位で取得するのにWin32APIを使用
'SYSTEMTIME構造体定義
Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
'関数定義
Public Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
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
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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