Attribute VB_Name = "TableSheet"
Option Explicit
Option Private Module

'====================================================================================================
'
' �e�[�u���V�[�g���W���[��
'
'====================================================================================================

'====================================================================================================
' �e�[�u���V�[�g�̍쐬
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ�̘A���z��
'    : tableDefinitions �e�[�u����`�̘A���z��
'====================================================================================================
Public Sub CreateTableSheet(tableSettings As Object, tableDefinitions As Object)
On Error GoTo Finally
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim columnRange As Variant
    Dim ws As Worksheet
    Dim xKey As Variant

    Dim requiredDic As Object
    Set requiredDic = CreateObject("Scripting.Dictionary")
    Call requiredDic.Add("1", "�K�{")

    Dim primaryKeyDic As Object
    Set primaryKeyDic = CreateObject("Scripting.Dictionary")
    Call primaryKeyDic.Add("1", "PK")

    Dim tmplSheet As Object
    Set tmplSheet = ThisWorkbook.Worksheets(cstSheetTemplate)
    tmplSheet.Visible = True

    ' �e�[�u���V�[�g���폜����
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> cstSheetMain And ws.Name <> cstSheetTemplate Then
            ThisWorkbook.Worksheets(ws.Name).Delete
        End If
    Next

    For Each xKey In tableSettings
        Set td = tableDefinitions(xKey)
        ' �e���v���[�g�V�[�g���R�s�[����
        tmplSheet.Copy Before:=tmplSheet
        ThisWorkbook.ActiveSheet.Name = td.TableName

        ' �R�s�[�����V�[�g�ɃJ������`����������
        If td.ColumnDefinitions.Count > 0 Then
            ReDim columnRange(1 To ColumnDefinitionRow.Max, 1 To td.ColumnDefinitions.Count)
            For Each cd In td.ColumnDefinitions
                columnRange(ColumnDefinitionRow.Comments, cd.ColumnId) = cd.Comments
                columnRange(ColumnDefinitionRow.ColumnName, cd.ColumnId) = cd.ColumnName
                columnRange(ColumnDefinitionRow.DataType, cd.ColumnId) = cd.DataType
                columnRange(ColumnDefinitionRow.DataLength, cd.ColumnId) = cd.DataLength
                columnRange(ColumnDefinitionRow.IsRequired, cd.ColumnId) = requiredDic(cd.IsRequired)
                columnRange(ColumnDefinitionRow.IsPrimaryKey, cd.ColumnId) = primaryKeyDic(cd.IsPrimaryKey)
            Next
            ThisWorkbook.Worksheets(td.TableName).Range(Cells(1, 1), Cells(ColumnDefinitionRow.Max, td.ColumnDefinitions.Count)) = columnRange
        End If
    Next
Finally:
    ThisWorkbook.Worksheets(cstSheetMain).Activate
    tmplSheet.Visible = False
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


'====================================================================================================
' �����f�[�^���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xTableName �e�[�u����
' OUT: �����f�[�^
'====================================================================================================
Public Function GetEntryData(xTableName As String) As EntryData
    Dim ed As EntryData
    Dim cd As ColumnDefinition
    Dim rowIndex As Long
    Dim dataCount As Long
    Dim columnRange As Range
    Dim cr As Range

    With ThisWorkbook.Worksheets(xTableName)
        Set ed = New EntryData
        ed.TableName = xTableName

        ' �J������`���X�g�̍쐬
        ed.ColumnDefinitions = New Collection
        dataCount = WorkSheetEx.GetColDataCount(xTableName, ColumnDefinitionRow.ColumnName, 1)
        Set columnRange = .Range(.Cells(1, 1), .Cells(ColumnDefinitionRow.Max, dataCount))
        For Each cr In columnRange.Columns
            Set cd = New ColumnDefinition
            cd.ColumnName = cr.Cells(ColumnDefinitionRow.ColumnName, 1)
            cd.DataType = cr.Cells(ColumnDefinitionRow.DataType, 1)
            cd.IsPrimaryKey = cr.Cells(ColumnDefinitionRow.IsPrimaryKey, 1)
            Call ed.ColumnDefinitions.Add(cd)
        Next

        ' ���R�[�h�͈͂̐ݒ�
        dataCount = WorkSheetEx.GetRowDataCount(xTableName, cstTableRecordBase, 1)
        ed.RecordRange = .Range(.Cells(cstTableRecordBase, 1), .Cells(cstTableRecordBase + dataCount - 1, ed.ColumnDefinitions.Count))

        Set GetEntryData = ed
    End With
End Function
