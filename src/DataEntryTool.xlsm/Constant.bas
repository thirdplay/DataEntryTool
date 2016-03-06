Attribute VB_Name = "Constant"
Option Explicit

'====================================================================================================
'
' 定数モジュール
'
'====================================================================================================
' データベース種類
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' 投入種類
Public Enum EntryType
    Register = 0        ' 登録
    Update              ' 更新
    Remove              ' 削除
End Enum
