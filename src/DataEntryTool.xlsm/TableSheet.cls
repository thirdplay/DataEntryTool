VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' テーブルシート
'
'====================================================================================================

'====================================================================================================
' テーブルシートの作成
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定の連装配列
'    : tableDefinitions テーブル定義の連装配列
'====================================================================================================
Public Sub CreateTableSheet(tableSettings As Dictionary, tableDefinitions As Dictionary)
On Error GoTo Finally
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim columnRange As Variant
    Dim ws As Worksheet
    Dim xKey As Variant

    Dim requiredDic As Dictionary
    Set requiredDic = New Dictionary
    Call requiredDic.Add("1", "必須")

    Dim primaryKeyDic As Dictionary
    Set primaryKeyDic = New Dictionary
    Call primaryKeyDic.Add("1", "PK")

    Dim tmplSheet As Worksheet
    Set tmplSheet = ThisWorkbook.Worksheets(cstSheetTemplate)
    tmplSheet.Visible = True

    ' テーブルシートを削除する
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> cstSheetMain And ws.Name <> cstSheetTemplate Then
            ThisWorkbook.Worksheets(ws.Name).Delete
        End If
    Next

    For Each xKey In tableSettings
        Call ApplicationEx.CheckEvents

        Set td = tableDefinitions(xKey)
        ' テンプレートシートをコピーする
        tmplSheet.Copy Before:=tmplSheet
        ThisWorkbook.ActiveSheet.Name = td.TableName

        ' コピーしたシートにカラム定義を書き込む
        If td.ColumnDefinitions.Count > 0 Then
            ReDim columnRange(1 To ColumnDefinitionRow.Max, 1 To td.ColumnDefinitions.Count)
            For Each cd In td.ColumnDefinitions
                columnRange(ColumnDefinitionRow.Comments, cd.ColumnId) = cd.Comments
                columnRange(ColumnDefinitionRow.ColumnName, cd.ColumnId) = cd.ColumnName
                columnRange(ColumnDefinitionRow.DataType, cd.ColumnId) = cd.DataType
                columnRange(ColumnDefinitionRow.DataLength, cd.ColumnId) = cd.DataLength
                columnRange(ColumnDefinitionRow.IsRequired, cd.ColumnId) = requiredDic(cd.IsRequired)
                columnRange(ColumnDefinitionRow.IsPrimaryKey, cd.ColumnId) = primaryKeyDic(cd.IsPrimaryKey)
            Next
            ThisWorkbook.Worksheets(td.TableName).Range(Cells(1, 1), Cells(ColumnDefinitionRow.Max, td.ColumnDefinitions.Count)) = columnRange
        End If
    Next
Finally:
    ThisWorkbook.Worksheets(cstSheetMain).Activate
    tmplSheet.Visible = False
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


'====================================================================================================
' 投入データを取得します
'----------------------------------------------------------------------------------------------------
' IN : xTableName テーブル名
' OUT: 投入データ
'====================================================================================================
Public Function GetEntryData(xTableName As String) As EntryData
    Dim ed As EntryData
    Dim cd As ColumnDefinition
    Dim rowIndex As Long
    Dim dataCnt As Long
    Dim columnRange As Range
    Dim cr As Range

    With ThisWorkbook.Worksheets(xTableName)
        Set ed = New EntryData
        ed.TableName = xTableName

        ' カラム定義リストの作成
        ed.ColumnDefinitions = New Collection
        dataCnt = WorkSheetEx.GetColDataCount(xTableName, ColumnDefinitionRow.ColumnName, 1)
        Set columnRange = .Range(.Cells(1, 1), .Cells(ColumnDefinitionRow.Max, dataCnt))
        For Each cr In columnRange.Columns
            Set cd = New ColumnDefinition
            cd.ColumnName = cr.Cells(ColumnDefinitionRow.ColumnName, 1)
            cd.DataType = cr.Cells(ColumnDefinitionRow.DataType, 1)
            cd.IsPrimaryKey = cr.Cells(ColumnDefinitionRow.IsPrimaryKey, 1)
            Call ed.ColumnDefinitions.Add(cd)
        Next

        ' レコード範囲の設定
        dataCnt = WorkSheetEx.GetRowDataCount(xTableName, cstTableRecordBase, 1)
        ed.RecordRange = .Range(.Cells(cstTableRecordBase, 1), .Cells(cstTableRecordBase + dataCnt - 1, ed.ColumnDefinitions.Count))

        Set GetEntryData = ed
    End With
End Function
