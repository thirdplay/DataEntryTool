Attribute VB_Name = "TableDefinition"
Option Explicit


'====================================================================================================
' テーブル定義を最新化します。
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

    MsgBox "テーブル定義再作成完了"
    Exit Sub
ErrHandler:
    Set control = Nothing
    MsgBox Err.Description
    Exit Sub
End Sub
