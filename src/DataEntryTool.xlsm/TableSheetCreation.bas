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
    Set tableDefinitions = Database.GetColumnDefinitions(tableSettings)
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
