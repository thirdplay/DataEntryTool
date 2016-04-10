Attribute VB_Name = "DataEntry"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�������W���[��
'
'====================================================================================================

'====================================================================================================
' �f�[�^��o�^���܂�
'====================================================================================================
Public Sub RegisterData()
    Call Execute(EntryType.Register)
End Sub


'====================================================================================================
' �f�[�^���X�V���܂�
'====================================================================================================
Public Sub UpdateData()
    Call Execute(EntryType.Update)
End Sub


'====================================================================================================
' �f�[�^���폜���܂�
'====================================================================================================
Public Sub RemoveData()
    Call Execute(EntryType.Remove)
End Sub


'====================================================================================================
' �f�[�^�����̎��s
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Private Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim operationDic As Object
    Dim tableSettings As Object
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim xKey As Variant
    Dim procCount As Long

    ' ������
    If Not Initialize Then
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
        Set ed = DataEntryModel.GetEntryData(xEntryType, ts.PhysicsName)

        ' �f�[�^�������s
        procCount = DataEntryModel.ExecuteDataEntry(ed)

        ' ���������̏�������
        Call TableSheet.WriteProcessingCount(ts, procCount)
    Next
Finally:
    ' �I����
    Call Finalize

    ' ���s���ʂ̕\��
    Set operationDic = GetOperationDic()
    Call ApplicationEx.ShowExecutionResult(Err.Number = 0, "�f�[�^" & operationDic(xEntryType))
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
    If Not Setting.CheckDataEntrySetting() Then
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


'====================================================================================================
' ������ނɑΉ�����������������i�[���鎫�����擾���܂�
'----------------------------------------------------------------------------------------------------
' OUT: ��������
'====================================================================================================
Private Function GetOperationDic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Call dic.Add(EntryType.Register, "�o�^")
    Call dic.Add(EntryType.Update, "�X�V")
    Call dic.Add(EntryType.Remove, "�폜")
    Set GetOperationDic = dic
End Function

