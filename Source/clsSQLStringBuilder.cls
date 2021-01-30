VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'SQL文字操作（主に結合）のためのクラス
'開始履歴、終了履歴を設定して、文字列連結は配列のJoinで行う方向

Private strSQL As String                '最後の結果のSQL文
Private arrstrField() As String         'フィールドの文字列配列
Private intFieldCount As Integer        'フィールドの数
Private intBulkCount As Integer         'バルク数
Private strStartRireki As String        'スタートの履歴
Private strEndRireki As String          'エンドの履歴
Private longStringLength As Long        '文字列の長さ
Private strTableName As String          'テーブル名
Private intCurrentValueCount As Integer '現在の処理中データの場所（バルクの時に使う）


Private Sub Class_Initialize()
    'クラス初期化動作
    strSQL = Empty
    intFieldCount = 0
    intCurrentValueCount = 0
    intBulkCount = 500                  'バルク数初期設定
    strStartRireki = Empty
    strEndRireki = Empty
    longStringLength = 0
End Sub

Private Sub Class_Terminate()
    'クラス破棄時に実行
    strSQL = Empty
    intFieldCount = 0
    strStartRireki = Empty
    strEndRireki = Empty
    longStringLength = 0
End Sub
'プロパティ設定
Property Let FieldArray(arrstrargField() As String)
    'フィールド名を配列の形で指定
    arrstrField = arrstrargField
End Property
Property Get FieldArray() As String()
    'フィールドの配列を返す
    FieldArray = arrstrField
End Property
Property Let BulkCount(intargBulkCount As Integer)
    'バルク数設定
    intBulkCount = intargBulkCount
End Property
Property Get BulkCount() As Integer
    'バルク数取得
    BulkCount = intBulkCount
End Property
Property Let StartRireki(strargStartRireki As String)
    'スタート履歴設定
    strStartRireki = strargStartRireki
End Property
Property Get StartRireki() As String
    'スタート履歴取得
    StartRireki = strStartRireki
End Property
Property Let EndRireki(strargEndRireki As String)
    'エンド履歴設定
    strEndRireki = strargEndRireki
End Property
Property Get EndRireki() As String
    'エンド履歴取得
    EndRireki = strEndRireki
End Property
Property Let TableName(strargTableName As String)
    'テーブル名設定
    strTableName = strargTableName
End Property
Property Get TableName() As String
    'テーブル名取得
    TableName = strTableName
End Property

'メソｄｄッド
Public Function CreateInsertSQL(Optional ByVal strargStartRireki As String, _
                                Optional ByVal strargEndRireki As String, _
                                Optional ByVal strargTableName As String) As Boolean
    'フィールド名は先にプロパティでセットしてほしいの・・・
    If Not strargStartRireki = "" Then
        Me.StartRireki = strargStartRireki
    End If
    If Not strargEndRireki = "" Then
        Me.EndRireki = strargEndRireki
    End If
    If Not strargTableName = "" Then
        Me.TableName = strargTableName
    End If
    
    If Me.StartRireki = "" Or _
        Me.EndRireki = "" Or _
        Me.TableName = "" Then
        MsgBox "CreateInsertSQL エラー：必須項目で入力されていない場所があります"
        CreateInsertSQL = False
        Exit Function
    End If
    'DBから履歴の桁数とか拾ってくる
    
End Function
