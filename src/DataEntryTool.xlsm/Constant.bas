Attribute VB_Name = "Constant"
Option Explicit
Option Private Module

'====================================================================================================
'
' 定数モジュール
'
'====================================================================================================
' 名前定義
Public Const cstDatabaseType = "DatabaseType"               ' データベース種類
Public Const cstServerName = "ServerName"                   ' サーバ名
Public Const cstPort = "Port"                               ' ポート
Public Const cstDatabaseName = "DatabaseName"               ' データベース名
Public Const cstUserId = "UserId"                           ' ユーザID
Public Const cstPassword = "Password"                       ' パスワード
Public Const cstLinefeedCode = "LinefeedCode"               ' 改行コード
Public Const cstDateFormat = "DateFormat"                   ' 日付書式
Public Const cstTableBase = "TableBase"                     ' テーブル一覧の基準セル

' データベース種類
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' 投入種類
Public Enum EntryType
    Register = 0        ' 登録
    Update              ' 更新
    Remove              ' 削除
End Enum
