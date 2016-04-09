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
    Dim dataCnt As Long

    dataCnt = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex + dataCnt, colIndex).Value <> ""
            dataCnt = dataCnt + 1
        Loop
    End With
    GetRowDataCount = dataCnt
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
    Dim dataCnt As Long

    dataCnt = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex, colIndex + dataCnt).Value <> ""
            dataCnt = dataCnt + 1
        Loop
    End With
    GetColDataCount = dataCnt
End Function

