Attribute VB_Name = "TableDefinitionModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�x�[�X�̃e�[�u�������Ƀe�[�u���V�[�g���쐬���郂�f��
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
'Private mDatabaseModel As DatabaseModel   ' �f�[�^�x�[�X���f��


'====================================================================================================
' �e�[�u���ݒ胊�X�g�̃e�[�u������DB����擾���A�ԋp���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ胊�X�g
' OUT: �e�[�u����`���X�g
'====================================================================================================
Public Function GetTableDefinitions(tableSettings As Collection) As Collection
    Dim ts As TableSetting
    Dim td As TableDefinition
    Dim list As Collection
    Dim xDatabaseModel As DatabaseModel

    Set xDatabaseModel = DatabaseModelFactory.Create()

    Set list = New Collection
    For Each ts In tableSettings
        Set td = New TableDefinition
        td.ColumnDefinitions = xDatabaseModel.GetColumnDefinitions(ts.PhysicsName)
        If td.ColumnDefinitions.Count = 0 Then
            Err.Raise 100, , "�e�[�u��[" & ts.PhysicsName & "]�̃J������`���擾�ł��܂���B"
        End If
        td.TableName = ts.PhysicsName
        Call list.Add(td)
    Next

    Set GetTableDefinitions = list
End Function
