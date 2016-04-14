Attribute VB_Name = "DataEntry"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�������W���[��
'
'====================================================================================================

'====================================================================================================
' �f�[�^�o�^
'====================================================================================================
Public Sub RegisterData()
    Call Execute(EntryType.Register)
End Sub


'====================================================================================================
' �f�[�^�X�V
'====================================================================================================
Public Sub UpdateData()
    Call Execute(EntryType.Update)
End Sub


'====================================================================================================
' �f�[�^�폜
'====================================================================================================
Public Sub DeleteData()
    Call Execute(EntryType.Delete)
End Sub


'====================================================================================================
' �f�[�^�����̎��s
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Private Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim tableSettings As Object
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim xKey As Variant
    Dim procCnt As Long
    Dim operationDic As Object
    Set operationDic = GetOperationDic()

    ' �}�N���N��
    Call ApplicationEx.StartupMacro(MacroType.DataEntry)

    ' ���������̃N���A
    Call DataEntrySheet.ClearProcessingCount

    ' �Ώۃe�[�u���ݒ�̎擾
    Set tableSettings = DataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "�f�[�^�����Ώۂ̃f�[�^������܂���B" & vbNewLine & vbNewLine & _
            "���L�菇�����{���ăf�[�^�����Ώۂ̃f�[�^��ݒ肵�Ă��������B" & vbNewLine & _
            "  �E�e�[�u���ꗗ�̃f�[�^�����Ώۗ�ɋ󕶎��ȊO�̒l��ݒ肷��B" & vbNewLine & _
            "  �E�f�[�^�����Ώۂ̃e�[�u���V�[�g�Ƀf�[�^����͂���B"
    End If

    ' �Ώۃe�[�u���ݒ��S�ď���
    For Each xKey In tableSettings
        Set ts = tableSettings(xKey)

        ' �Ώۃe�[�u���̃e�[�u���f�[�^�̎擾
        Set ed = TableSheet.GetEntryData(ts.PhysicsName)

        ' �f�[�^�������s
        procCnt = DataEntryModel.ExecuteDataEntry(xEntryType, ed)

        ' ���������̏�������
        Call DataEntrySheet.WriteProcessingCount(ts, procCnt)
    Next
Finally:
    ' �}�N����~
    Call ApplicationEx.ShutdownMacro

    ' ���s���ʂ̕\��
    Call ApplicationEx.ShowExecutionResult("�f�[�^" & operationDic(xEntryType))
End Sub


'====================================================================================================
' ������ނɑΉ�����������������i�[�����A���z����擾���܂�
'----------------------------------------------------------------------------------------------------
' OUT: �A���z��
'====================================================================================================
Private Function GetOperationDic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Call dic.Add(EntryType.Register, "�o�^")
    Call dic.Add(EntryType.Update, "�X�V")
    Call dic.Add(EntryType.Delete, "�폜")
    Set GetOperationDic = dic
End Function
