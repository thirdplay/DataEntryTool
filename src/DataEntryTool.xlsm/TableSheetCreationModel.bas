Attribute VB_Name = "TableSheetCreationModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �e�[�u���V�[�g�쐬���f���̃��W���[��
'
'====================================================================================================

'====================================================================================================
' �e�[�u���ݒ胊�X�g�̃e�[�u������DB����擾���A�ԋp���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ胊�X�g
' OUT: �e�[�u����`���X�g
'====================================================================================================
Public Function GetTableDefinitions(tableSettings As Object) As Collection
    Dim ts As TableSetting
    Dim td As TableDefinition
    Dim list As Collection
    Dim xKey As Variant

    Set list = New Collection
    For Each xKey In tableSettings
        Set ts = tableSettings(xKey)
        Set td = New TableDefinition
        td.ColumnDefinitions = Database.GetColumnDefinitions(ts.PhysicsName)
        If td.ColumnDefinitions.Count = 0 Then
            Err.Raise 1000, , "�e�[�u��[" & ts.PhysicsName & "]�̃J������`���擾�ł��܂���B"
        End If
        td.TableName = ts.PhysicsName
        Call list.Add(td)
    Next

    Set GetTableDefinitions = list
End Function


