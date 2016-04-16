Attribute VB_Name = "DataEntryTool"
Option Explicit
'Option Private Module

'====================================================================================================
'
' データ投入ツールのモジュール
'
'====================================================================================================

'====================================================================================================
' XLSMファイルオープンイベント
'====================================================================================================
Private Sub Auto_Open()
On Error GoTo ErrHandler
    ' 参照設定
    Call ThisWorkbook.VBProject.References.AddFromGuid(cstGuidAdodb, 6, 1)
    Call ThisWorkbook.VBProject.References.AddFromGuid(cstGuidScripting, 1, 0)
    Exit Sub
ErrHandler:
    If Err.Number = ErrNumber.AlreadyReferenceConfigured Then
        Resume Next
    Else
        MsgBox "Error Number : " & Err.Number & vbNewLine & Err.Description
    End If
End Sub


'====================================================================================================
' テーブルシート作成
'====================================================================================================
Public Sub CreateTableSheet()
    Dim controller As TableSheetController

    Set controller = New TableSheetController
    Call controller.Execute
End Sub
