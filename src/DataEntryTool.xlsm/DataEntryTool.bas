Attribute VB_Name = "DataEntryTool"
Option Explicit


'====================================================================================================
' �e�[�u���V�[�g���쐬���܂�
'====================================================================================================
Public Sub CreateTableSheet()
On Error GoTo Finally
    Dim control As DataEntryControl
    Dim start As Single
    Dim finish As Single
    start = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .CreateTableSheet
    End With

Finally:
    finish = Timer
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox "�e�[�u���V�[�g�̍쐬�Ɏ��s���܂���" & vbNewLine & Err.Description
    Else
        Debug.Print "��������:" & (finish - start)
        MsgBox "�e�[�u���V�[�g�̍쐬���������܂���" & ":" & (finish - start)
    End If
End Sub


'====================================================================================================
' �f�[�^��o�^���܂�
'====================================================================================================
Public Sub RegisterData()
    Call ExecuteEntryData(EntryType.Register)
End Sub


'====================================================================================================
' �f�[�^���X�V���܂�
'====================================================================================================
Public Sub UpdateData()
    Call ExecuteEntryData(EntryType.Update)
End Sub


'====================================================================================================
' �f�[�^���폜���܂�
'====================================================================================================
Public Sub RemoveData()
    Call ExecuteEntryData(EntryType.Remove)
End Sub


'====================================================================================================
' �f�[�^�����̎��s
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Private Sub ExecuteEntryData(xEntryType As EntryType)
On Error GoTo Finally
    Dim operationStr As String
    Dim control As DataEntryControl
    Dim start As Single
    Dim finish As Single
    start = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .EntryData(xEntryType)
    End With

Finally:
    finish = Timer
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    operationStr = GetOperationStr(xEntryType)
    If Err.Number <> 0 Then
        MsgBox "�f�[�^" & operationStr & "�Ɏ��s���܂���" & vbNewLine & Err.Description
    Else
        Debug.Print "��������:" & (finish - start)
        MsgBox "�f�[�^" & operationStr & "���������܂���" & ":" & (finish - start)
    End If
End Sub


'====================================================================================================
' ������ނɑΉ����鑀�앶������擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
' OUT: ���앶����
'====================================================================================================
Private Function GetOperationStr(xEntryType As EntryType)
    Select Case xEntryType
        Case EntryType.Register
            GetOperationStr = "�o�^"
        Case EntryType.Register
            GetOperationStr = "�X�V"
        Case EntryType.Register
            GetOperationStr = "�폜"
    End Select
End Function

