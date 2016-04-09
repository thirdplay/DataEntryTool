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
    Dim tableSettings As Collection
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim procCount As Long

    ' ��ʕ`��̗}��
    Call ApplicationEx.SuppressScreenDrawing(True)
    ' �ݒ胂�W���[���̍\��
    Call Setting.Setup
    If Not Setting.CheckDataEntrySetting() Then
        Exit Sub
    End If

    ' ���������̃N���A
    Call ProcessCountModel.ClearProcessingCount

    ' �Ώۃe�[�u���ݒ�̎擾
    Set tableSettings = TableSettingModel.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "�f�[�^�����Ώۂ̃f�[�^������܂���B" & vbNewLine & vbNewLine & _
            "���L�菇�����{���ăf�[�^�����Ώۂ̃f�[�^��ݒ肵�Ă��������B" & vbNewLine & _
            "  �E�e�[�u���ꗗ�̃f�[�^�����Ώۗ�ɋ󕶎��ȊO�̒l��ݒ肷��B" & vbNewLine & _
            "  �E�f�[�^�����Ώۂ̃e�[�u���V�[�g�Ƀf�[�^����͂���B"
            
        Exit Sub
    End If

    ' �Ώۃe�[�u���ݒ��S�ď���
    For Each ts In tableSettings
        ' �Ώۃe�[�u���̃e�[�u���f�[�^�̎擾
        Set ed = EntryDataModel.GetEntryData(xEntryType, ts.PhysicsName)

        ' �f�[�^�������s
        procCount = DataEntryModel.ExecuteDataEntry(ed)

        ' ���������̏�������
        Call ProcessCountModel.WriteProcessingCount(ts, procCount)
    Next
Finally:
    ' ��ʕ`��̗}������
    Call ApplicationEx.SuppressScreenDrawing(False)

    ' ���s���ʂ̕\��
    Set operationDic = GetOperationDic
    If Err.Number <> 0 Then
        MsgBoxEx.Error "�f�[�^" & operationDic(xEntryType) & "�Ɏ��s���܂���" & vbNewLine & Err.Description
    Else
        MsgBox "�f�[�^" & operationDic(xEntryType) & "���������܂���"
    End If
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

