Attribute VB_Name = "TableDefinitionModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' データベースのテーブルを元にテーブルシートを作成するモデル
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
'Private mDatabaseModel As DatabaseModel   ' データベースモデル


'====================================================================================================
' テーブル設定リストのテーブル情報をDBから取得し、返却します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定リスト
' OUT: テーブル定義リスト
'====================================================================================================
Public Function GetTableDefinitions(tableSettings As Collection) As Collection
    Dim ts As TableSetting
    Dim td As TableDefinition
    Dim list As Collection
    Dim xDatabaseModel As DatabaseModel

    Set xDatabaseModel = DatabaseModelFactory.Create()

    Set list = New Collection
    For Each ts In tableSettings
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
