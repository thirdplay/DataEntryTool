VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntryData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' 投入データ
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
' Insertクエリ
Private Const cstInsertQuery = "INSERT INTO ${tableName} (${columns}) VALUES (${values})"
' Updateクエリ
Private Const cstUpdateQuery = "UPDATE ${tableName} SET ${set} WHERE ${where}"
' Deleteクエリ
Private Const cstDeleteQuery = "DELETE FROM ${tableName} WHERE ${where}"
' 等価比較条件
Private Const cstEqualComparisonCriteria = "${column} = ${value}"
' 区切り文字
Private Const cstDelimiter = ","
' Where句の区切り文字
Private Const cstWherePhraseDelimiter = " AND "


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mTableName As String                ' テーブル名
Private mColumnDefinitions As Collection    ' カラム定義リスト
Private mRecordRange As Range               ' レコード範囲


'====================================================================================================
' テーブル名の取得/設定
'====================================================================================================
Public Property Get TableName() As String
    TableName = mTableName
End Property
Public Property Let TableName(TableName As String)
    mTableName = TableName
End Property


'====================================================================================================
' カラム定義リストの取得/設定
'====================================================================================================
Public Property Get ColumnDefinitions() As Collection
    Set ColumnDefinitions = mColumnDefinitions
End Property
Public Property Let ColumnDefinitions(ColumnDefinitions As Collection)
    Set mColumnDefinitions = ColumnDefinitions
End Property


'====================================================================================================
' レコード範囲の取得/設定
'====================================================================================================
Public Property Get RecordRange() As Range
    Set RecordRange = mRecordRange
End Property
Public Property Let RecordRange(RecordRange As Range)
    Set mRecordRange = RecordRange
End Property


'====================================================================================================
' Insertクエリを生成します
'----------------------------------------------------------------------------------------------------
' OUT: クエリリスト
'====================================================================================================
Public Function MakeInsertQueries() As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To Me.RecordRange.Rows.Count
        query = cstInsertQuery
        query = Replace(query, "${tableName}", Me.TableName)
        query = Replace(query, "${columns}", GetColumnPhrase(Me.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(Me.ColumnDefinitions, Me.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeInsertQueries = queries
End Function


'====================================================================================================
' Updateクエリを生成します
'----------------------------------------------------------------------------------------------------
'' OUT: クエリリスト
'====================================================================================================
Public Function MakeUpdateQueries() As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To Me.RecordRange.Rows.Count
        query = cstUpdateQuery
        query = Replace(query, "${tableName}", Me.TableName)
        query = Replace(query, "${set}", GetSetPhrase(Me.ColumnDefinitions, Me.RecordRange.Rows(i)))
        query = Replace(query, "${where}", GetWherePhrase(Me.ColumnDefinitions, Me.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeUpdateQueries = queries
End Function


'====================================================================================================
' Deleteクエリを生成します
'----------------------------------------------------------------------------------------------------
' OUT: クエリリスト
'====================================================================================================
Public Function MakeDeleteQueries() As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To Me.RecordRange.Rows.Count
        query = cstDeleteQuery
        query = Replace(query, "${tableName}", Me.TableName)
        query = Replace(query, "${where}", GetWherePhrase(Me.ColumnDefinitions, Me.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeDeleteQueries = queries
End Function


'====================================================================================================
' Insert文のColumn句を取得します
'----------------------------------------------------------------------------------------------------
' IN : xColumnDefinitions カラム定義リスト
' OUT: Column句
'====================================================================================================
Private Function GetColumnPhrase(xColumnDefinitions As Collection) As String
    Dim result As String
    Dim cd As ColumnDefinition

    For Each cd In xColumnDefinitions
        result = result & cd.ColumnName & cstDelimiter
    Next
    GetColumnPhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' Insert文のValue句を取得します
'----------------------------------------------------------------------------------------------------
' IN : cd カラム定義リスト
'    : record レコード
' OUT: Value句
'====================================================================================================
Private Function GetValuePhrase(cd As Collection, record As Range) As String
    Dim result As String
    Dim i As Long

    For i = 1 To record.Columns.Count
        result = result & Database.GetItemValue(record.Cells(1, i), cd(i).DataType) & cstDelimiter
    Next
    GetValuePhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' Update文のSet句を取得します
'----------------------------------------------------------------------------------------------------
' IN : cd カラム定義リスト
'    : record レコード
' OUT: Set句
'====================================================================================================
Private Function GetSetPhrase(cd As Collection, record As Range) As String
    Dim result As String
    Dim i As Long

    For i = 1 To record.Columns.Count
        If cd(i).IsPrimaryKey = "" Then
            result = result & cstEqualComparisonCriteria
            result = Replace(result, "${column}", cd(i).ColumnName)
            result = Replace(result, "${value}", Database.GetItemValue(record.Cells(1, i), cd(i).DataType))
            result = result & cstDelimiter
        End If
    Next
    GetSetPhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' Where句を取得します
'----------------------------------------------------------------------------------------------------
' IN : cd カラム定義リスト
'    : record レコード
' OUT: Value句
'====================================================================================================
Private Function GetWherePhrase(cd As Collection, record As Range) As String
    Dim result As String
    Dim i As Long

    For i = 1 To record.Columns.Count
        If cd(i).IsPrimaryKey <> "" Then
            result = result & cstEqualComparisonCriteria
            result = Replace(result, "${column}", cd(i).ColumnName)
            result = Replace(result, "${value}", Database.GetItemValue(record.Cells(1, i), cd(i).DataType))
            result = result & cstWherePhraseDelimiter
        End If
    Next
    GetWherePhrase = Left(result, Len(result) - Len(cstWherePhraseDelimiter))
End Function
