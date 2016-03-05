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
        .CreateTableSheet
    End With

Finally:
    Finish = Timer
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description
    Else
        Debug.Print "��������:" & (Finish - Start)
        MsgBox "�e�[�u����`�č쐬����" & ":" & (Finish - Start)
    End If
End Sub
