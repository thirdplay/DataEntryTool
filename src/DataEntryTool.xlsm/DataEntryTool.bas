Attribute VB_Name = "DataEntryTool"
Option Explicit


'====================================================================================================
' �e�[�u����`���ŐV�����܂��B
'====================================================================================================
Public Sub RecreateTableDefinition()
On Error GoTo Finally
    Dim control As DataEntryControl

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        .RecreateTableDefinition
    End With

Finally:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description
    Else
        MsgBox "�e�[�u����`�č쐬����"
    End If
End Sub
