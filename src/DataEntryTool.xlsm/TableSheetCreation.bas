Attribute VB_Name = "TableSheetCreation"
Option Explicit
Option Private Module

'====================================================================================================
'
' テーブルシート作成モジュール
'
'====================================================================================================

'====================================================================================================
' テーブルシート作成の実行
'====================================================================================================
Public Sub Execute()
On Error GoTo Finally
    Dim tableSettings As Object
    Dim tableDefinitions As Object

    ' マクロ起動
    Call ApplicationEx.StartupMacro(MacroType.Database)

    ' テーブル設定リストの取得
    Set tableSettings = DataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "作成対象のテーブルがありません。" & vbNewLine & "テーブル一覧にテーブル物理名を入力してください。"
    End If

    ' テーブル設定リストを元に、テーブル定義リストを取得
    Set tableDefinitions = GetColumnDefinitions(tableSettings)
    ' テーブル定義リストを元に、テーブルシートを作成する
    Call TableSheet.CreateTableSheet(tableSettings, tableDefinitions)
    ' テーブル設定にハイパーリンクを設定する
    Call DataEntrySheet.SetHyperlink(tableSettings)

Finally:
    ' マクロ停止
    Call ApplicationEx.ShutdownMacro

    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult("テーブルシートの作成")
End Sub


'====================================================================================================
' カラム定義リストを取得します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定の連想配列
' OUT: カラム定義リスト
'====================================================================================================
Private Function GetColumnDefinitions(tableSettings As Object) As Object
    Dim rs As Object
    Dim tableNames As Object
    Dim tableNameInStr As String
    Dim xTableName As Variant
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim dic As Object

    ' テーブル名を取得し、連想配列を作成する
    Set rs = Database.GetTableName()
    Set tableNames = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        Call tableNames.Add(rs("table_name").Value, rs("table_name").Value)
        rs.MoveNext
    Loop

    ' テーブル名の存在チェック
    For Each xTableName In tableSettings
        If Not tableNames.Exists(xTableName) Then
            Err.Raise ErrNumber.Error, , "テーブル[" & xTableName & "]のカラム定義が取得できません。"
        End If
    Next

    ' カラム定義の取得
    Set rs = Database.GetColumnDefinition(tableSettings)

    ' テーブル定義の連装配列を作成する
    Set dic = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        ' テーブル定義生成
        If Not dic.Exists(rs("table_name").Value) Then
            Set td = New TableDefinition
            td.TableName = rs("table_name").Value
            td.ColumnDefinitions = New Collection
            Call dic.Add(td.TableName, td)
        End If

        ' カラム定義生成
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = rs("column_id").Value
            .ColumnName = rs("column_name").Value
            .Comments = rs("comments").Value
            .DataType = rs("data_type").Value
            .DataLength = rs("data_length").Value
            .IsRequired = rs("is_required").Value
            .IsPrimaryKey = rs("is_primary_key").Value
        End With
        Call dic(rs("table_name").Value).ColumnDefinitions.Add(cd)
        rs.MoveNext
    Loop
    Set GetColumnDefinitions = dic
End Function
