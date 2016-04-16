VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSheetLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' テーブルシートのロジック
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDataEntryDao As DataEntryDao


'====================================================================================================
' コンストラクタ
'====================================================================================================
Public Sub Class_Initialize()
    Set mDataEntryDao = DataEntryDaoFactory.GetDataEntryDao
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Public Sub Class_Terminate()
    Set mDataEntryDao = Nothing
End Sub


'====================================================================================================
' カラム定義リストを取得します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定の連想配列
' OUT: カラム定義リスト
'====================================================================================================
Public Function GetColumnDefinitions(tableSettings As Dictionary) As Dictionary
    Dim rs As Object
    Dim tableNames As Dictionary
    Dim tableNameInStr As String
    Dim xTableName As Variant
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim dic As Dictionary

    ' テーブル名を取得し、連想配列を作成する
    Set rs = mDataEntryDao.GetTableName()
    Set tableNames = New Dictionary
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
    Set rs = mDataEntryDao.GetColumnDefinition(tableSettings)

    ' テーブル定義の連装配列を作成する
    Set dic = New Dictionary
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
