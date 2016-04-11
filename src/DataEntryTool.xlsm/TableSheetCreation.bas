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

    ' ������
    If Not Initialize() Then
        Exit Sub
    End If

    ' �e�[�u���ݒ胊�X�g�̎擾
    Set tableSettings = DataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "�쐬�Ώۂ̃e�[�u��������܂���B", "�e�[�u���ꗗ�Ƀe�[�u������������͂��Ă��������B"
        Exit Sub
    End If

    ' �e�[�u���ݒ胊�X�g�����ɁA�e�[�u����`���X�g���擾
    Set tableDefinitions = Database.GetColumnDefinitions(tableSettings)
    ' �e�[�u����`���X�g�����ɁA�e�[�u���V�[�g���쐬����
    Call TableSheet.CreateTableSheet(tableDefinitions)
    ' �e�[�u���ݒ�Ƀn�C�p�[�����N��ݒ肷��
    Call DataEntrySheet.SetHyperlink(tableSettings)

Finally:
    ' �I����
    Call Finalize

    ' ���s���ʂ̕\��
    Call ApplicationEx.ShowExecutionResult("�e�[�u���V�[�g�̍쐬")
End Sub


'====================================================================================================
' ������
'----------------------------------------------------------------------------------------------------
' OUT: True:�����AFalse:���s
'====================================================================================================
Private Function Initialize() As Boolean
    ' ��ʕ`��̗}��
    Call ApplicationEx.SuppressScreenDrawing(True)

    ' �ݒ胂�W���[���̍\��
    Call Setting.Setup
    If Not Setting.CheckDbSetting() Then
        Initialize = False
        Exit Function
    End If

    ' �f�[�^�x�[�X�ڑ�
    Call Database.Connect
    Initialize = True
End Function


'====================================================================================================
' �I����
'====================================================================================================
Private Sub Finalize()
    ' �f�[�^�x�[�X�ؒf
    Call Database.Disconnect

    ' ��ʕ`��̗}������
    Call ApplicationEx.SuppressScreenDrawing(False)
End Sub
