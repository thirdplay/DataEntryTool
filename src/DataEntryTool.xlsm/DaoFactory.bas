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
Public Function Create(Setting As Setting) As IDataEntryDao
    If Setting.DatabaseType = cstDatabaseTypeOracle Then
        Set Create = New DataEntryOracleDao
    ElseIf Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
        Set Create = New DataEntryPostgreSqlDao
    End If
End Function
