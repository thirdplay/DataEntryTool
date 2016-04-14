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
    Dim tableSettings As Object
    Dim tableDefinitions As Object

    ' �}�N���N��
    Call ApplicationEx.StartupMacro(MacroType.Database)

    ' �e�[�u���ݒ胊�X�g�̎擾
    Set tableSettings = DataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "�쐬�Ώۂ̃e�[�u��������܂���B" & vbNewLine & "�e�[�u���ꗗ�Ƀe�[�u������������͂��Ă��������B"
    End If

    ' �e�[�u���ݒ胊�X�g�����ɁA�e�[�u����`���X�g���擾
    Set tableDefinitions = GetColumnDefinitions(tableSettings)
    ' �e�[�u����`���X�g�����ɁA�e�[�u���V�[�g���쐬����
    Call TableSheet.CreateTableSheet(tableSettings, tableDefinitions)
    ' �e�[�u���ݒ�Ƀn�C�p�[�����N��ݒ肷��
    Call DataEntrySheet.SetHyperlink(tableSettings)

Finally:
    ' �}�N����~
    Call ApplicationEx.ShutdownMacro

    ' ���s���ʂ̕\��
    Call ApplicationEx.ShowExecutionResult("�e�[�u���V�[�g�̍쐬")
End Sub


'====================================================================================================
' �J������`���X�g���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ�̘A�z�z��
' OUT: �J������`���X�g
'====================================================================================================
Private Function GetColumnDefinitions(tableSettings As Object) As Object
    Dim rs As Object
    Dim tableNames As Object
    Dim tableNameInStr As String
    Dim xTableName As Variant
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim dic As Object

    ' �e�[�u�������擾���A�A�z�z����쐬����
    Set rs = Database.GetTableName()
    Set tableNames = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        Call tableNames.Add(rs("table_name").Value, rs("table_name").Value)
        rs.MoveNext
    Loop

    ' �e�[�u�����̑��݃`�F�b�N
    For Each xTableName In tableSettings
        If Not tableNames.Exists(xTableName) Then
            Err.Raise ErrNumber.Error, , "�e�[�u��[" & xTableName & "]�̃J������`���擾�ł��܂���B"
        End If
    Next

    ' �J������`�̎擾
    Set rs = Database.GetColumnDefinition(tableSettings)

    ' �e�[�u����`�̘A���z����쐬����
    Set dic = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        ' �e�[�u����`����
        If Not dic.Exists(rs("table_name").Value) Then
            Set td = New TableDefinition
            td.TableName = rs("table_name").Value
            td.ColumnDefinitions = New Collection
            Call dic.Add(td.TableName, td)
        End If

        ' �J������`����
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = rs("column_id").Value
            .ColumnName = rs("column_name").Value
            .Comments = rs("comments").Value
            .DataType = rs("data_type").Value
            .DataLength = rs("data_length").Value
            .IsRequired = rs("is_required").Value
            .IsPrimaryKey = rs("is_primary_key").Value
        End With
        Call dic(rs("table_name").Value).ColumnDefinitions.Add(cd)
        rs.MoveNext
    Loop
    Set GetColumnDefinitions = dic
End Function
