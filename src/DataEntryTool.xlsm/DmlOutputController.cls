VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DmlOutputController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' DML�o�͂̃R���g���[���[
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mBaseController As BaseController       ' ��{�R���g���[���[
Private mDataEntrySheet As DataEntrySheet       ' �f�[�^�����V�[�g
Private mTableSheet As TableSheet               ' �e�[�u���V�[�g
Private mDmlOutputLogic As DmlOutputLogic       ' DML�o�̓��W�b�N
Private mProcessName As String                  ' ������


'====================================================================================================
' ���s
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Public Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim tableSettings As Dictionary
    Dim xOutputDir As String
    Dim xKey As Variant
    Dim ts As TableSetting
    Dim ed As EntryData

    ' ������
    Call Initialize(xEntryType)

    ' �o�͐�̎擾
    xOutputDir = Setting.OutputDir
    If xOutputDir = "" Then
        Err.Raise ErrNumber.Abort
    End If
    If Dir(xOutputDir, vbDirectory) = "" Then
        Err.Raise ErrNumber.Warning, , "�o�͐�[" & xOutputDir & "]�����݂��܂���B"
    End If
    If Right(xOutputDir, 1) <> "\" Then
        xOutputDir = xOutputDir & "\"
    End If

    ' �Ώۃe�[�u���ݒ�̎擾
    Set tableSettings = mDataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "�f�[�^�����Ώۂ̃f�[�^������܂���B" & vbNewLine & vbNewLine & _
            "���L�菇�����{���ăf�[�^�����Ώۂ̃f�[�^��ݒ肵�Ă��������B" & vbNewLine & _
            "  �E�e�[�u���ꗗ�̃f�[�^�����Ώۗ�ɋ󕶎��ȊO�̒l��ݒ肷��B" & vbNewLine & _
            "  �E�f�[�^�����Ώۂ̃e�[�u���V�[�g�Ƀf�[�^����͂���B"
    End If

    ' �Ώۃe�[�u���ݒ��S�ď���
    For Each xKey In tableSettings
        ' �Ώۃe�[�u���̃e�[�u���f�[�^�̎擾
        Set ts = tableSettings(xKey)
        Set ed = mTableSheet.GetEntryData(ts.PhysicsName)

        ' DML�o��
        Call mDmlOutputLogic.OutputDml(xOutputDir, xEntryType, ed)
    Next

Finally:
    ' �I����
    Call Finalize
End Sub


'====================================================================================================
' ������
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Private Sub Initialize(xEntryType As EntryType)
    ' �N��
    Set mBaseController = New BaseController
    Call mBaseController.Startup(SettingType.DataEntry)

    ' �C���X�^���X����
    Set mDataEntrySheet = New DataEntrySheet
    Set mTableSheet = New TableSheet
    Set mDmlOutputLogic = New DmlOutputLogic
    mProcessName = GetProcessName(xEntryType)
End Sub


'====================================================================================================
' �I����
'====================================================================================================
Private Sub Finalize()
    ' �C���X�^���X���
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing
    Set mDmlOutputLogic = Nothing

    ' ��~
    Call mBaseController.Shutdown

    ' ���s���ʂ̕\��
    Call mBaseController.ShowExecutionResult(mProcessName)
    Set mBaseController = Nothing
End Sub


'====================================================================================================
' ���������擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
' OUT: ������
'====================================================================================================
Private Function GetProcessName(xEntryType As EntryType)
    Select Case xEntryType
        Case EntryType.Register
            GetProcessName = "�o�^DML�o��"
        Case EntryType.Update
            GetProcessName = "�X�VDML�o��"
        Case EntryType.Delete
            GetProcessName = "�폜DML�o��"
    End Select
End Function