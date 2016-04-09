Attribute VB_Name = "WorkSheetEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' ワークシートの拡張モジュール
'
'====================================================================================================

'====================================================================================================
' 行データ数を取得します
'----------------------------------------------------------------------------------------------------
' IN : sheetName シート名
'    : rowIndex 行インデックス
'    : colIndex 列インデックス
' OUT: 行データ数
'====================================================================================================
Public Function GetRowDataCount(sheetName As String, rowIndex As Long, colIndex As Long) As Long
    Dim count As Long

    count = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex + count, colIndex).Value <> ""
            count = count + 1
        Loop
    End With
    GetRowDataCount = count
End Function


'====================================================================================================
' 列データ数を取得します
'----------------------------------------------------------------------------------------------------
' IN : sheetName シート名
'    : rowIndex 行インデックス
'    : colIndex 列インデックス
' OUT: 列データ数
'====================================================================================================
Public Function GetColDataCount(sheetName As String, rowIndex As Long, colIndex As Long) As Long
    Dim count As Long

    count = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex, colIndex + count).Value <> ""
            count = count + 1
        Loop
    End With
    GetColDataCount = count
End Function

