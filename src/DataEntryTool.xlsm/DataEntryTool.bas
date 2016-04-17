Attribute VB_Name = "DataEntryTool"
Option Explicit
'Option Private Module

'====================================================================================================
'
' データ投入ツールのモジュール
'
'====================================================================================================

'====================================================================================================
' Excelファイルオープンイベント
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

'====================================================================================================
' データ登録
'====================================================================================================
Public Sub RegisterData()
    Call ExecuteDataEntry(EntryType.Register)
End Sub


'====================================================================================================
' データ更新
'====================================================================================================
Public Sub UpdateData()
    Call ExecuteDataEntry(EntryType.Update)
End Sub


'====================================================================================================
' データ削除
'====================================================================================================
Public Sub DeleteData()
    Call ExecuteDataEntry(EntryType.Delete)
End Sub


'====================================================================================================
' データ投入の実行
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
'====================================================================================================
Private Sub ExecuteDataEntry(xEntryType As EntryType)
    Dim controller As DataEntryController

    Set controller = New DataEntryController
    Call controller.Execute(xEntryType)
End Sub
