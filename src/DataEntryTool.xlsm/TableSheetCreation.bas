Attribute VB_Name = "TableSheetCreation"
Option Explicit
Option Private Module

'====================================================================================================
'
' �e�[�u���V�[�g�쐬���W���[��
'
'====================================================================================================

'====================================================================================================
' �e�[�u���V�[�g�쐬�̎��s
'====================================================================================================
Public Sub Execute()
On Error GoTo Finally
    Dim tableSettings As Collection
    Dim tableDefinitions As Collection

    ' ��ʕ`��̗}��
    Call ApplicationEx.SuppressScreenDrawing(True)
    ' �ݒ胂�W���[���̍\��
    Call Setting.Setup
    If Not Setting.CheckDbSetting() Then
        Exit Sub
    End If

    ' �e�[�u���ݒ胊�X�g�̎擾
    Set tableSettings = TableSettingModel.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "�쐬�Ώۂ̃e�[�u��������܂���B" & vbNewLine & "�e�[�u���ꗗ�Ƀe�[�u������������͂��Ă��������B"
        Exit Sub
    End If

    ' �e�[�u���ݒ胊�X�g�����ɁA�e�[�u����`���X�g���擾
    Set tableDefinitions = TableDefinitionModel.GetTableDefinitions(tableSettings)
    ' �e�[�u����`���X�g�����ɁA�e�[�u���V�[�g���쐬����
    Call CreateTableSheet(tableDefinitions)

Finally:
    ' ��ʕ`��̗}������
    Call ApplicationEx.SuppressScreenDrawing(False)

    ' ���s���ʂ̕\��
    If Err.Number <> 0 Then
        MsgBoxEx.Error "�e�[�u���V�[�g�̍쐬�Ɏ��s���܂���" & vbNewLine & Err.Description
    Else
        MsgBox "�e�[�u���V�[�g�̍쐬���������܂���"
    End If
End Sub


'====================================================================================================
' �e�[�u���V�[�g�̍쐬
'----------------------------------------------------------------------------------------------------
' IN : tableDefinitions �e�[�u����`���X�g
'====================================================================================================
Private Sub CreateTableSheet(tableDefinitions As Collection)
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
