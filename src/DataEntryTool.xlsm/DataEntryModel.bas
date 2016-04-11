Attribute VB_Name = "DataEntryModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�������f���̃��W���[��
'
'====================================================================================================

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


'====================================================================================================
' �f�[�^���������s���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: ��������
'====================================================================================================
Public Function ExecuteDataEntry(xEntryData As EntryData) As Long
On Error GoTo ErrHandler
    Dim i As Long
    Dim query As String
    Dim procCnt As Long

    ' �g�����U�N�V�����J�n
    Call Database.BeginTrans

    ' �N�G������
    query = Database.MakeDataEntryQuery(xEntryData)
    ' �N�G�����s
    Call Database.ExecuteQuery(query)

    ' �R�~�b�g
    Call Database.CommitTrans

    ExecuteDataEntry = xEntryData.RecordRange.Rows.Count
    Exit Function
ErrHandler:
    ' ���[���o�b�N
    Call Database.RollbackTrans
    Err.Raise Err.Number, Err.Source, _
        "[�������]" & vbNewLine & _
        "�e�[�u����:" & xEntryData.TableName & vbNewLine & _
        "[�G���[���e]" & vbNewLine & _
        Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function
