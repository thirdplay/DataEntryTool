Attribute VB_Name = "TableDefinition"
Option Explicit


'====================================================================================================
' �e�[�u����`���ŐV�����܂��B
'====================================================================================================
Public Sub Recreate()
On Error GoTo ErrHandler
    Dim control As TableDefinitionControl

    Set control = New TableDefinitionControl
    With control
        .View = New DataEntryView
        .Model = New TableDefinitionModel
        .Execute
    End With
    Set control = Nothing

    MsgBox "�e�[�u����`�č쐬����"
    Exit Sub
ErrHandler:
    Set control = Nothing
    MsgBox Err.Description
    Exit Sub
End Sub
