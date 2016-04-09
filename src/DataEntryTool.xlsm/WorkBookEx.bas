Attribute VB_Name = "WorkBookEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' ワークブックの拡張モジュール
'
'====================================================================================================

'====================================================================================================
' シートが存在するか判定します
'----------------------------------------------------------------------------------------------------
' IN : sheetName シート名
' OUT: 存在する場合はtrue、それ以外はfalse
'====================================================================================================
Public Function ExistsSheet(sheetName As String) As Boolean
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.Name = sheetName Then
            ExistsSheet = True
            Exit Function
        End If
    Next
    ExistsSheet = False
End Function


