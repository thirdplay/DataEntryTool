Attribute VB_Name = "Constant"
Option Explicit
Option Private Module

'====================================================================================================
'
' 定数モジュール
'
'====================================================================================================
' データベース種類
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' 改行コード
Public Const cstLinefeedCodeCRLF = "CRLF"                   ' CRLF
Public Const cstLinefeedCodeLF = "LF"                       ' LF

' 投入種類
Public Enum EntryType
    Register = 0        ' 登録
    Update              ' 更新
    Remove              ' 削除
End Enum
