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
' IN : tableDefinitions �e�[�u����`���X�g
'====================================================================================================
Public Sub CreateTableSheet(tableDefinitions As Collection)
On Error GoTo Finally
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim columnRange As Variant
    Dim ws As Worksheet

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

    For Each td In tableDefinitions
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
' �����������N���A���܂�
'====================================================================================================
Public Sub ClearProcessingCount()
    Dim baseRowIndex As Long
    Dim dataCount As Long

    With ThisWorkbook.Worksheets(cstSheetMain)
        ' �e�[�u���ݒ萔�̎擾
        baseRowIndex = .Range(cstTableBase).Row + 1
        dataCount = WorkSheetEx.GetRowDataCount(cstSheetMain, baseRowIndex, TableSettingCol.PhysicsName)

        ' ���������̃N���A
        Call .Range(.Cells(baseRowIndex, TableSettingCol.ProcessCount), .Cells(baseRowIndex + dataCount, TableSettingCol.ProcessCount)).ClearContents
    End With
End Sub


'====================================================================================================
' ������������������
'----------------------------------------------------------------------------------------------------
' IN : xTableSetting �e�[�u���ݒ�
'    : procCount ��������
'====================================================================================================
Public Sub WriteProcessingCount(xTableSetting As TableSetting, procCount As Long)
    With ThisWorkbook.Worksheets(cstSheetMain)
        .Cells(xTableSetting.Row, TableSettingCol.ProcessCount).Value = procCount
    End With
End Sub

