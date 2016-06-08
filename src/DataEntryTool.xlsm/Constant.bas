Attribute VB_Name = "Constant"
Option Explicit
Option Private Module

'====================================================================================================
'
' 定数モジュール
'
'====================================================================================================
' シート名
Public Const cstSheetMain = "データ投入ツール"              ' メインシート
Public Const cstSheetTemplate = "テンプレート"              ' テンプレートシート

' 名前定義
Public Const cstDatabaseType = "DatabaseType"               ' データベース種別
Public Const cstServerName = "ServerName"                   ' サーバ名
Public Const cstServerNameLabel = "ServerNameLabel"         ' サーバ名ラベル
Public Const cstUserId = "UserId"                           ' ユーザID
Public Const cstPassword = "Password"                       ' パスワード
Public Const cstPort = "Port"                               ' ポート
Public Const cstDatabaseName = "DatabaseName"               ' データベース名
Public Const cstLinefeedCode = "LinefeedCode"               ' 改行コード
Public Const cstDateFormat = "DateFormat"                   ' 日付書式
Public Const cstTimestampFormat = "TimestampFormat"         ' タイムスタンプ書式
Public Const cstOutputDir = "OutputDir"                     ' 出力先
Public Const cstTableBase = "TableBase"                     ' テーブル一覧の基準セル

' データベース種別
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' サーバ名ラベル
Public Const cstServerNameLabelOracle = "接続識別子"
Public Const cstServerNameLabelPostgreSQL = "ホスト名"

' 改行コード
Public Const cstLinefeedCodeCRLF = "CRLF"                   ' CRLF
Public Const cstLinefeedCodeLF = "LF"                       ' LF

' エラー番号
Public Enum ErrNumber
    Error = 1000        ' ユーザ定義のエラー
    Warning = 2000      ' ユーザ定義の警告
    Abort = 3000        ' ユーザ定義の中断
End Enum

' 設定種別
Public Enum SettingType
    None = &H0                          ' なし
    Database = &H1                      ' データベース
    DataEntry = &H2 Or Database         ' データ投入
End Enum

' 投入種別
Public Enum EntryType
    Register = 0        ' 登録
    Update              ' 更新
    Delete              ' 削除
End Enum

' テーブル設定の列参照値
Public Enum TableSettingCol
    PhysicsName = 1         ' 物理名
    LogicalName             ' 論理名
    DataEntryTarget         ' データ操作対象
    ProcessCount            ' 処理件数
    Max = ProcessCount
End Enum

' カラム定義の行参照値
Public Enum ColumnDefinitionRow
    Comments = 1            ' コメント
    ColumnName              ' 列名
    DataType                ' データ型
    DataLength              ' データ長
    IsRequired              ' 必須指定
    IsPrimaryKey            ' 主キー指定
    Max = IsPrimaryKey
End Enum
' テーブルレコードの開始行
Public Const cstTableRecordBase = ColumnDefinitionRow.Max + 1

' カーソルロケーション
Public Const adUseClient = 3
' コマンドを非同期に実行することを示す
Public Const adAsyncExecute = &H10
