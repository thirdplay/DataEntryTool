Attribute VB_Name = "DaoFactory"
Option Explicit

'====================================================================================================
'
' DAO�������W���[��
'
'====================================================================================================


'====================================================================================================
' �f�[�^�x�[�X��ʂɑΉ�����DAO�𐶐����ĕԋp���܂�
'====================================================================================================
Public Function Create(Setting As Setting) As IDataEntryDao
    If Setting.DatabaseType = cstDatabaseTypeOracle Then
        Set Create = New OracleDao
    ElseIf Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
        Set Create = New PostgreSqlDao
    End If
End Function
