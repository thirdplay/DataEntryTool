Attribute VB_Name = "TableDefinition"
Option Explicit

'====================================================================================================
' �e�[�u����`���ŐV�����܂��B
'====================================================================================================
Public Sub Recreate()
    Dim control As TableDefinitionControl
    
    Set control = New TableDefinitionControl
    With control
        .View = New DataEntryView
        .Model = New TableDefinitionModel
        .Execute
    End With
    Set control = Nothing
    
    MsgBox "�e�[�u����`�č쐬����"
End Sub
