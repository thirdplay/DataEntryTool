Attribute VB_Name = "TableSettingModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' テーブル設定モデルのモジュール
'
'====================================================================================================

'====================================================================================================
' データ投入対象のテーブル設定リストを返却します
'----------------------------------------------------------------------------------------------------
' IN : isEntryTarget 投入対象フラグ(True:対象のみ,False:全て)
' OUT: テーブル設定リスト
'====================================================================================================
Public Function GetTableSettings(isEntryTarget As Boolean) As Object
    Dim rowIndex As Long
    Dim ts As TableSetting
    Dim isTarget As Boolean
    Dim xTableName As String
    Dim dataCount As Long
    Dim tableRange As Range
    Dim tr As Range
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    With ThisWorkbook.Worksheets(cstSheetMain)
        rowIndex = .Range(cstTableBase).Row + 1
        dataCount = WorkSheetEx.GetRowDataCount(cstSheetMain, rowIndex, TableSettingCol.PhysicsName)
        If dataCount > 0 Then
            Set tableRange = .Range(.Cells(rowIndex, TableSettingCol.PhysicsName), .Cells(rowIndex + dataCount - 1, TableSettingCol.Max))

            For Each tr In tableRange.Rows
                ' 投入対象フラグがtrueの場合、投入対象外のテーブルは処理しない
                xTableName = tr.Cells(1, TableSettingCol.PhysicsName).Value
                isTarget = True
                If isEntryTarget Then
                    If tr.Cells(1, TableSettingCol.DataEntryTarget).Value = "" Then
                        isTarget = False
                    ElseIf Not WorkBookEx.ExistsSheet(xTableName) Then
                        Err.Raise 100, , "テーブル[" & xTableName & "]のシートが存在しません。" & vbNewLine & "テーブルシート作成を行い、テーブルシートを作成してください。"
                    ElseIf ThisWorkbook.Worksheets(xTableName).Cells(cstTableRecordBase, 1).Value = "" Then
                        isTarget = False
                    End If
                End If

                If isTarget Then
                    If dic.Exists(tr.Cells(1, TableSettingCol.PhysicsName).Value) Then
                        Err.Raise 100, , "テーブル[" & xTableName & "]が重複しています。"
                    End If
                    Set ts = New TableSetting
                    ts.Row = tr.Row
                    ts.PhysicsName = tr.Cells(1, TableSettingCol.PhysicsName).Value
                    ts.LogicalName = tr.Cells(1, TableSettingCol.LogicalName).Value
                    ts.DataEntryTarget = tr.Cells(1, TableSettingCol.DataEntryTarget).Value
                    Call dic.Add(ts.PhysicsName, ts)
                End If
            Next
        End If
    End With
    Set GetTableSettings = dic
End Function


