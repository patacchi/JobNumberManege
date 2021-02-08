Attribute VB_Name = "PublicConst"
Option Explicit
'Option Base 1
'office 2013導入により、mdbからaccdb形式に移行
'DBをSQLite3に移行  2021_01_10 Pataccchi
Public Const constDatabasePath              As String = "Database_sqlite3"      'データベースディレクトリ
Public Const constJobNumberDBname           As String = "JobNumberDB.sqlite3"   'ジョブ番号情報DBのファイル名
Public Const Field_Initialdate              As String = "InitialInputDate"      '各テーブル共通、初回入力時刻
Public Const Field_Update                   As String = "UpdateDate"            '各テーブル共通、最終更新時刻
Public Const Field_BarcordeNumber           As String = "BarcodeNumber"         'テーブル共通、トレサIDバーコードデータ
Public Const Field_LocalInput               As String = "LocalInput"            'テーブル共通、ローカル入力時 1（ローカル・リモートバックアップ判定）
Public Const Field_RemoteInput              As String = "RemoteInput"           'テーブル共通、リモート入力時 1（こっちが本体になるはず）
'ジョブ番号管理DBテーブル・フィールド名定義
'機種別情報格納テーブル定義
Public Const Table_Kishu                    As String = "T_Kishu"               '機種別情報格納テーブル名
Public Const Kishu_Header                   As String = "KishuHeader"           '機種判別用ヘッダ情報フィールド（重複不可）
Public Const Kishu_KishuName                As String = "KishuName"             '機種名フィールド P70664A
Public Const Kishu_KishuNickname            As String = "KishuNickName"         '機種通称名 マスター
Public Const Kishu_TotalKeta                As String = "TotalRirekiketa"       '総桁数フィールド（多分20しかないと思う）
Public Const Kishu_RenbanKetasuu            As String = "RenbanKetasuu"         '連番桁数フィールド
Public Const Kishu_Mai_Per_Sheet            As String = "Mai_Per_Sheet"         '1シートあたりの枚数
Public Const Kishu_Barcord_Read_Number      As String = "Barcord_Read_Number"   'バーコード読み取り数
'ログ情報格納テーブル定義
'ログは基本的にトリガーで入力する
Public Const Table_Log                      As String = "T_Log"                 'ログ情報テーブル名
Public Const Log_ActionType                 As String = "ActionType"            'INSERT,UPDATE,DELETEのいずれかが入る
Public Const Log_Table                      As String = "TableName"             'テーブル名
Public Const Log_StartRireki                As String = "StartRireki"           '(INSERT)開始履歴
Public Const Log_Maisuu                     As String = "Maisuu"                '(INSERT)枚数
Public Const Log_JobNumber                  As String = "JobNumber"             '(INSERT) T_JobData_ ジョブ番号
Public Const Log_RirekiHeader               As String = "RirekiHeader"          '(INSERT) T_JobData_ 履歴ヘッダ
Public Const Log_BarcordNumber              As String = "BarcodeNumber"         '(INSERT) T_Barcode_ バーコード番号
Public Const Log_SQL                        As String = "SQL"                   '(UPDATE,DELETE)この場合はSQLのみ格納
'ジョブ・履歴情報テーブル定義
Public Const Table_JobDataPri               As String = "T_JobData_"            'ジョブ履歴テーブル名前半部分、実際はこの後に機種名が連結されてテーブル名となる
Public Const Job_Number                     As String = "JobNumber"             'ジョブ番号フィールド名
Public Const Job_RirekiHeader               As String = "RirekiHeader"          '履歴ヘッダフィールド名
Public Const Job_RirekiNumber               As String = "RirekiNumber"          '履歴の連番部分（Longで格納）
Public Const Job_Rireki                     As String = "Rireki"                'ヘッダ+履歴連番（作成するか要検討）
Public Const Job_KanbanChr                  As String = "KanbanChr"             'カンバン分割文字列
Public Const Job_ProductDate                As String = "ProductDate"           '製造予定日（計画引き当てに使う）
'こっちは全部機種別にテーブルを分ける
'通常バーコード記録テーブル
Public Const Table_Barcodepri               As String = "T_Barcode_"            '機種別バーコード入力情報テーブル、実際は後半に機種名が連結される
Public Const Laser_Rireki                   As String = "LaserRirekiNumber"       'レーザーの履歴番号格納フィールド（Longで格納）、重複不可
'再印字等バーコード記録テーブル
Public Const Table_Retrypri                 As String = "T_Retry_"              '機種別再印字バーコード履歴格納テーブル、実際は後半に機種名が連結される
Public Const Retry_Rireki                   As String = "LaserRetryRireki"      '再印字の履歴フィールド名（Longで格納）、再印字は履歴重複OK
Public Const Retry_Reason                   As String = "RetryReason"           '再印字理由フィールド
'フィールド追加用SQL定型文
Public Const strAddField1_NextTableName     As String = "ALTER TABLE """        '追加の最初、この次にテーブル名が入る
Public Const strAddField2_NextFieldName     As String = """ ADD COLUMN """      '二番目、この次にフィールド名が入る
Public Const strADDField3_Text_Last         As String = """ TEXT;"              '最後、ただしTEXT型の場合
Public Const strAddFiels3_Numeric           As String = """ NUMERIC;"           '数値の場合の最後
'インデックス追加用SQL定型文
Public Const strIndex1_NextTable            As String = "CREATE INDEX IF NOT EXISTS ""ixJob_"
Public Const strIndex2_NextTable            As String = """ ON """
Public Const strIndes3_Field1               As String = """ ("""
Public Const strIndex4_FieldNext            As String = """ ASC ,"""            '複数フィールドに対して実行する場合は、以後これの繰り返し
Public Const strIndex5_Last                 As String = """ ASC);"
'テーブル追加用SQL定型文
Public Const strAddTable1_NextTable         As String = "CREATE TABLE IF NOT EXISTS """ 'テーブル追加用定型文ここから
Public Const strAddTable2_Field1            As String = """ ("""                'フィールドの最初だけこいつを使う
Public Const strAddTable_TEXT_Next_Field    As String = """ TEXT,"""            '紛らわしいけど、「前」がText型の場合こっちを使う、次にフィールド名が続く
Public Const strAddTable_NUMELIC_Next_Field As String = """ NUMERIC,"""         '「前」がNumericの場合はこっち
Public Const strAddTable_Text_Last          As String = """ TEXT);"             'メンドウなので、最後はTextで終わらせて・・・
Public Const strAddTable_Numeric_Last       As String = """ NUMERIC);"          '一応数値型で終わるやつも
'シート情報格納関係定数
Public Const constMaisuu_Label              As String = "Maisuu"                '履歴枚数（単独セル参照）名前定義
Public Const constRirekiFromLabel           As String = "Rireki_From"           '履歴From（単独セル参照）名前定義
Public Const constRirekiToLabel             As String = "Rireki_To"             '履歴To（単独セル参照）名前定義
Public Const constMaxRirekiKetasuu          As Byte = 20                        '履歴桁数のMax値
Public Const constDefaultArraySize          As Long = 6000                      'DBからの結果セットの配列の初期上限
Public Const constAddArraySize              As Long = 2000                      '配列確保行数が足りなくなった場合の1回で増量する分
'エラーコード定義（もう使わないかも・・・
Public Const errcNone                       As Integer = 0                      '正常終了
Public Const errcDBAlreadyExistValue        As Integer = -2                      '既に同じ値がDB上に有る場合
Public Const errcDBFileNotFound             As Integer = -4                      'DBファイル見つからないよぅ
Public Const errcDBFieldNotFont             As Integer = -8                      'DBで指定されたフィールドが見つからない
Public Const errcxlNameNotFound             As Integer = -16                     'Excelで名前定義が見つからない
Public Const errcxlDataNothing              As Integer = -32                     'ExcelでデータNothing
Public Const errcOthers                     As Integer = -16384                  'その他エラー
'機種情報を格納する構造体
Public Type typKishuInfo
    KishuHeader As String
    KishuName As String
    KishuNickName As String
    TotalRirekiketa As Byte
    RenbanKetasuu As Byte
End Type
Public Type typMaisuuRireki
    From As String
    To As String
End Type
Public arrKishuInfoGlobal() As typKishuInfo
'いんすーとする時のフィールド定義をもうここでハードコーディングしちゃう・・・
'テーブルが増えるたびに記述すること・・・
'どうやら配列は定数に出来ないようなので、SQLBuilderのコンストラクタ内で初期化する
Public arrFieldList_JobData() As String                                         'JobDataテーブルのフィールド定義
Public arrFieldList_Barcode() As String                                         'Barcodeテーブルのフィールド定義
Public arrFieldList_Retry() As String                                           'Retryテーブルのフィールド定義
Public oldMaisuData() As typMaisuuRireki
Public newMaisuData() As typMaisuuRireki
Public strRegistRireki As String                                                '機種登録時履歴、フォーム間の受け渡しに使う
Public strQRZuban As String                                                     '指示書QRコード読み取り時の図番格納、主に機種登録で使う
Public boolRegistOK As Boolean                                                  '機種登録が成功したらTrueフラグを立てる
Public boolNoTableKishuRecord As Boolean                                        '機種テーブルにデータが存在しない場合True、初期のみ