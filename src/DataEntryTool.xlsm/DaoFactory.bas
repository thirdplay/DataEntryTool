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
Public Function Create(databaseType As String) As IDataEntryDao
    If databaseType = cstDatabaseTypeOracle Then
        Set Create = New OracleDao
    ElseIf databaseType = cstDatabaseTypePostgreSQL Then
        Set Create = New PostgreSqlDao
    End If
End Function
