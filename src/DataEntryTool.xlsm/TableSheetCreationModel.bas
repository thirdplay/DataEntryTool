Attribute VB_Name = "TableSheetCreationModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' テーブルシート作成モデルのモジュール
'
'====================================================================================================

'====================================================================================================
' テーブル設定リストのテーブル情報をDBから取得し、返却します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定リスト
' OUT: テーブル定義リスト
'====================================================================================================
Public Function GetTableDefinitions(tableSettings As Object) As Collection
    Dim ts As TableSetting
    Dim td As TableDefinition
    Dim list As Collection
    Dim xKey As Variant
    Dim xDatabaseModel As DatabaseModel
    Set xDatabaseModel = DatabaseModelFactory.Create()

    Set list = New Collection
    For Each xKey In tableSettings
        Set ts = tableSettings(xKey)
        Set td = New TableDefinition
        td.ColumnDefinitions = xDatabaseModel.GetColumnDefinitions(ts.PhysicsName)
        If td.ColumnDefinitions.Count = 0 Then
            Err.Raise 100, , "テーブル[" & ts.PhysicsName & "]のカラム定義が取得できません。"
        End If
        td.TableName = ts.PhysicsName
        Call list.Add(td)
    Next

    Set GetTableDefinitions = list
End Function


'====================================================================================================
' テーブルシートの作成
'----------------------------------------------------------------------------------------------------
' IN : tableDefinitions テーブル定義リスト
'====================================================================================================
Public Sub CreateTableSheet(tableDefinitions As Collection)
On Error GoTo Finally
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim columnRange As Variant
    Dim ws As Worksheet

    Dim requiredDic As Object
    Set requiredDic = CreateObject("Scripting.Dictionary")
    Call requiredDic.Add("1", "必須")
    
    Dim primaryKeyDic As Object
    Set primaryKeyDic = CreateObject("Scripting.Dictionary")
    Call primaryKeyDic.Add("1", "PK")

    Dim tmplSheet As Object
    Set tmplSheet = ThisWorkbook.Worksheets(cstSheetTemplate)
    tmplSheet.Visible = True

    ' テーブルシートを削除する
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> cstSheetMain And ws.Name <> cstSheetTemplate Then
            ThisWorkbook.Worksheets(ws.Name).Delete
        End If
    Next

    For Each td In tableDefinitions
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


