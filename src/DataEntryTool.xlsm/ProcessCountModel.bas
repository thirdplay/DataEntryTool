Attribute VB_Name = "ProcessCountModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' 処理件数モデルのモジュール
'
'====================================================================================================

'====================================================================================================
' 処理件数をクリアします
'====================================================================================================
Public Sub ClearProcessingCount()
    Dim baseRowIndex As Long
    Dim dataCount As Long

    With ThisWorkbook.Worksheets(cstSheetMain)
        ' テーブル設定数の取得
        baseRowIndex = .Range(cstTableBase).Row + 1
        dataCount = WorkSheetEx.GetRowDataCount(cstSheetMain, baseRowIndex, TableSettingCol.PhysicsName)

        ' 処理件数のクリア
        Call .Range(.Cells(baseRowIndex, TableSettingCol.ProcessCount), .Cells(baseRowIndex + dataCount, TableSettingCol.ProcessCount)).ClearContents
    End With
End Sub


'====================================================================================================
' 処理件数を書き込む
'----------------------------------------------------------------------------------------------------
' IN : xTableSetting テーブル設定
'    : procCount 処理件数
'====================================================================================================
Public Sub WriteProcessingCount(xTableSetting As TableSetting, procCount As Long)
    With ThisWorkbook.Worksheets(cstSheetMain)
        .Cells(xTableSetting.Row, TableSettingCol.ProcessCount).Value = procCount
    End With
End Sub

