Attribute VB_Name = "DatabaseModelFactory"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�x�[�X���f���̃t�@�N�g���[
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Dim mDatabaseModel As DatabaseModel     ' �f�[�^�x�[�X���f��


'====================================================================================================
' �f�[�^�x�[�X���f���𐶐����܂�
'====================================================================================================
Public Function Create() As Object
    If mDatabaseModel Is Nothing Then
        Set mDatabaseModel = new DatabaseModel
    End If
    Set Create = mDatabaseModel
End Function

