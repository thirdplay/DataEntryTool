Attribute VB_Name = "TableSettingModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �e�[�u���ݒ胂�f���̃��W���[��
'
'====================================================================================================

'====================================================================================================
' �f�[�^�����Ώۂ̃e�[�u���ݒ胊�X�g��ԋp���܂�
'----------------------------------------------------------------------------------------------------
' IN : isEntryTarget �����Ώۃt���O(True:�Ώۂ̂�,False:�S��)
' OUT: �e�[�u���ݒ胊�X�g
'====================================================================================================
Public Function GetTableSettings(isEntryTarget As Boolean) As Collection
    Dim rowIndex As Long
    Dim ts As TableSetting
    Dim isTarget As Boolean
    Dim tableName As String
    Dim list As Collection
    Set list = New Collection

    With ThisWorkbook.Worksheets(cstSheetMain)
        rowIndex = .Range(cstTableBase).Row + 1
        Do While .Cells(rowIndex, TableSettingCol.PhysicsName).Value <> ""
            ' �����Ώۃt���O��true�̏ꍇ�A�����ΏۊO�̃e�[�u���͏������Ȃ�
            tableName = .Cells(rowIndex, TableSettingCol.PhysicsName).Value
            isTarget = True
            If isEntryTarget Then
                If .Cells(rowIndex, TableSettingCol.DataEntryTarget).Value = "" Then
                    isTarget = False
                ElseIf Not WorkBookEx.ExistsSheet(tableName) Then
                    isTarget = False
                ElseIf ThisWorkbook.Worksheets(tableName).Cells(cstTableRecordBase, 1).Value = "" Then
                    isTarget = False
                End If
            End If

            If isTarget Then
                Set ts = New TableSetting
                ts.Row = rowIndex
                ts.PhysicsName = .Cells(rowIndex, TableSettingCol.PhysicsName).Value
                ts.LogicalName = .Cells(rowIndex, TableSettingCol.LogicalName).Value
                ts.DataEntryTarget = .Cells(rowIndex, TableSettingCol.DataEntryTarget).Value
                Call list.Add(ts)
            End If
            rowIndex = rowIndex + 1
        Loop
    End With
    Set GetTableSettings = list
End Function


