Attribute VB_Name = "DaoFactory"
Option Explicit

'====================================================================================================
'
' DAO�������W���[��
'
'====================================================================================================


'====================================================================================================
' �f�[�^�x�[�X��ʂɑΉ�����DAO�𐶐����ĕԋp���܂��B
'====================================================================================================
Public Function Create(Setting As Setting) As IDataEntryDao
    If Setting.DatabaseType = cstDatabaseTypeOracle Then
        Set Create = New DataEntryOracleDao
    ElseIf Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
        Set Create = New DataEntryPostgreSqlDao
    End If
End Function
