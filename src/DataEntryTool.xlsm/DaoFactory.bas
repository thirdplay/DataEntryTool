Attribute VB_Name = "DaoFactory"
Option Explicit

'====================================================================================================
'
' DAO生成モジュール
'
'====================================================================================================

'====================================================================================================
' データベース種別に対応したDAOを生成して返却します。
'====================================================================================================
Public Function Create(databaseType As String) As IDataEntryDao
    If databaseType = cstDatabaseTypeOracle Then
        Set Create = New OracleDao
    ElseIf databaseType = cstDatabaseTypePostgreSQL Then
        Set Create = New PostgreSqlDao
    End If
End Function
