Attribute VB_Name = "DataEntry"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�������W���[��
'
'====================================================================================================

'====================================================================================================
' �f�[�^�����̎��s
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Private Sub Execute()
On Error GoTo Finally
    Dim tableSettings As Object
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim xKey As Variant
    Dim procCnt As Long

    ' �}�N���N��
    If Not ApplicationEx.StartupMacro(MacroType.DataEntry) Then
        Exit Sub
    End If

    ' ���������̃N���A
    Call TableSheet.ClearProcessingCount

    ' �Ώۃe�[�u���ݒ�̎擾
    Set tableSettings = DataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "�f�[�^�����Ώۂ̃f�[�^������܂���B", _
            "���L�菇�����{���ăf�[�^�����Ώۂ̃f�[�^��ݒ肵�Ă��������B" & vbNewLine & _
            "  �E�e�[�u���ꗗ�̃f�[�^�����Ώۗ�ɋ󕶎��ȊO�̒l��ݒ肷��B" & vbNewLine & _
            "  �E�f�[�^�����Ώۂ̃e�[�u���V�[�g�Ƀf�[�^����͂���B"
        Exit Sub
    End If

    ' �Ώۃe�[�u���ݒ��S�ď���
    For Each xKey In tableSettings
        Set ts = tableSettings(xKey)

        ' �Ώۃe�[�u���̃e�[�u���f�[�^�̎擾
        Set ed = DataEntryModel.GetEntryData(ts.PhysicsName)

        ' �f�[�^�������s
        procCnt = DataEntryModel.ExecuteDataEntry(ed)

        ' ���������̏�������
        Call TableSheet.WriteProcessingCount(ts, procCnt)
    Next
Finally:
    ' �}�N����~
    Call ApplicationEx.ShutdownMacro

    ' ���s���ʂ̕\��
    Call ApplicationEx.ShowExecutionResult("�f�[�^����")
End Sub
