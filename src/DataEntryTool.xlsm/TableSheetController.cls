VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSheetController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �e�[�u���V�[�g�̃R���g���[��
'
'====================================================================================================

'====================================================================================================
' �萔
'====================================================================================================
Private Const cstProcessName = "�e�[�u���V�[�g�쐬"


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDataEntrySheet As DataEntrySheet       ' �f�[�^�����V�[�g
Private mTableSheet As TableSheet               ' �e�[�u���V�[�g
Private mTableSheetLogic As TableSheetLogic     ' �e�[�u���V�[�g���W�b�N


'====================================================================================================
' ���s
'====================================================================================================
Public Sub Execute()
On Error GoTo Finally
    Dim tableSettings As Dictionary
    Dim tableDefinitions As Dictionary

    ' ������
    Call Initialize

    ' �e�[�u���ݒ�̎擾
    Set tableSettings = mDataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "�쐬�Ώۂ̃e�[�u��������܂���B" & vbNewLine & "�e�[�u���ꗗ�Ƀe�[�u������������͂��Ă��������B"
    End If

    ' �v���O���X�o�[�\��
    Call StatusBar.ShowProgress(cstProcessName, 100, 10)

    ' �e�[�u����`�̎擾
    Set tableDefinitions = mTableSheetLogic.GetColumnDefinitions(tableSettings)
    Call StatusBar.IncreaseProgress(50)

    ' �e�[�u���V�[�g�̍쐬
    Call mTableSheet.CreateTableSheet(tableSettings, tableDefinitions)
    Call StatusBar.IncreaseProgress(40)

    ' �n�C�p�[�����N�̐ݒ�
    Call mDataEntrySheet.SetHyperlink(tableSettings)
    Call StatusBar.IncreaseProgress(10)

Finally:
    ' �I����
    Call Finalize
    ' ���s���ʂ̕\��
    Call ApplicationEx.ShowExecutionResult(cstProcessName)
    ' �X�e�[�^�X�o�[��\��
    Call StatusBar.Hide
End Sub


'====================================================================================================
' ������
'====================================================================================================
Private Sub Initialize()
    ' �}�N���N��
    Call ApplicationEx.StartupMacro(SettingType.Database)

    ' �C���X�^���X����
    Set mDataEntrySheet = New DataEntrySheet
    Set mTableSheet = New TableSheet
    Set mTableSheetLogic = New TableSheetLogic
End Sub


'====================================================================================================
' �I����
'====================================================================================================
Private Sub Finalize()
    ' �C���X�^���X���
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing
    Set mTableSheetLogic = Nothing

    ' �}�N����~
    Call ApplicationEx.ShutdownMacro
End Sub