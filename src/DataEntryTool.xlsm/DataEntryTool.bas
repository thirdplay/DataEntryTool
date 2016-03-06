Attribute VB_Name = "DataEntryTool"
Option Explicit


'====================================================================================================
' �e�[�u���V�[�g���쐬���܂�
'====================================================================================================
Public Sub CreateTableSheet()
On Error GoTo Finally
    Dim control As DataEntryControl
    Dim Start As Single
    Dim Finish As Single
    Start = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .CreateTableSheet
    End With

Finally:
    Finish = Timer
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox "�e�[�u���V�[�g�̍쐬�Ɏ��s���܂���" & vbNewLine & Err.Description
    Else
        Debug.Print "��������:" & (Finish - Start)
        MsgBox "�e�[�u���V�[�g�̍쐬���������܂���" & ":" & (Finish - Start)
    End If
End Sub


'====================================================================================================
' �f�[�^��o�^���܂�
'====================================================================================================
Public Sub RegisterData()
On Error GoTo Finally
    Dim control As DataEntryControl

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .EntryData(EntryType.Register)
    End With

Finally:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox "�f�[�^�̓o�^�Ɏ��s���܂���" & vbNewLine & Err.Description
    Else
        MsgBox "�f�[�^�̓o�^���������܂���"
    End If
End Sub


'====================================================================================================
' �f�[�^���X�V���܂�
'====================================================================================================
Public Sub UpdateData()
    MsgBox "UpdateData"
End Sub


'====================================================================================================
' �f�[�^���폜���܂�
'====================================================================================================
Public Sub RemoveData()
    MsgBox "RemoveData"
End Sub


