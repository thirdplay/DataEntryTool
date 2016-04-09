Attribute VB_Name = "EntryDataModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �����f�[�^���f���̃��W���[��
'
'====================================================================================================

'====================================================================================================
' �����f�[�^���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'    : xTableName �e�[�u����
' OUT: �����f�[�^
'====================================================================================================
Public Function GetEntryData(xEntryType As EntryType, xTableName As String) As EntryData
    Dim ed As EntryData
    Dim cd As ColumnDefinition
    Dim rowIndex As Long
    Dim dataCount As Long
    Dim columnRange As Range
    Dim i As Long

    With ThisWorkbook.Worksheets(xTableName)
        Set ed = New EntryData
        ed.EntryType = xEntryType
        ed.TableName = xTableName

        ' �J������`���X�g�̍쐬
        ed.ColumnDefinitions = New Collection
        dataCount = WorkSheetEx.GetColDataCount(xTableName, ColumnDefinitionRow.ColumnName, 1)
        Set columnRange = .Range(.Cells(1, 1), .Cells(ColumnDefinitionRow.Max, dataCount))
        For i = 1 To columnRange.Columns.Count
            Set cd = New ColumnDefinition
            cd.ColumnName = columnRange.Cells(ColumnDefinitionRow.ColumnName, i)
            cd.DataType = columnRange.Cells(ColumnDefinitionRow.DataType, i)
            cd.IsPrimaryKey = columnRange.Cells(ColumnDefinitionRow.IsPrimaryKey, i)
            Call ed.ColumnDefinitions.Add(cd)
        Next

        ' ���R�[�h�͈͂̐ݒ�
        dataCount = WorkSheetEx.GetRowDataCount(xTableName, cstTableRecordBase, 1)
        ed.RecordRange = .Range(.Cells(cstTableRecordBase, 1), .Cells(cstTableRecordBase + dataCount - 1, ed.ColumnDefinitions.Count))

        Set GetEntryData = ed
    End With
End Function


