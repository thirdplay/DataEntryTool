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
Public Const cstTableBase = "TableBase"                     ' テーブル一覧の基準セル

' データベース種類
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' 改行コード
Public Const cstLinefeedCodeCRLF = "CRLF"                   ' CRLF
Public Const cstLinefeedCodeLF = "LF"                       ' LF

' マクロ種別
Public Enum MacroType
    Database = &H01                     ' データベース
    DataEntry = &H02 Or Database        ' データ投入
    Setting = &H04                      ' 設定
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
