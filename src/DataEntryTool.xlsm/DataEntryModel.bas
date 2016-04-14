Attribute VB_Name = "DataEntryModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' データ投入モデルのモジュール
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
' データ投入を実行します
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'    : xEntryData 投入データ
' OUT: 処理件数
'====================================================================================================
Public Function ExecuteDataEntry(xEntryType As EntryType, xEntryData As EntryData) As Long
On Error GoTo ErrHandler
    Dim i As Long
    Dim queries As Collection
    Dim query As String
    Dim procCnt As Long

    ' トランザクション開始
    Call Database.BeginTrans

    ' データ投入
    Set queries = MakeDataEntryQueries(xEntryType, xEntryData)
    For i = 1 To queries.Count
        procCnt = procCnt + Database.ExecuteQuery (queries(i))
    Next

    ' コミット
    Call Database.CommitTrans

    ExecuteDataEntry = procCnt
    Exit Function
ErrHandler:
    ' ロールバック
    Call Database.RollbackTrans
    Err.Raise Err.Number, Err.Source, _
        "[投入情報]" & vbNewLine & _
        "テーブル名:" & xEntryData.TableName & vbNewLine & _
        "行数:" & (cstTableRecordBase + i - 1) & vbNewLine & _
        "[エラー内容]" & vbNewLine & _
        Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function

'====================================================================================================
' クエリを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'    : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Public Function MakeDataEntryQueries(xEntryType As EntryType, xEntryData As EntryData) As Collection
    Select Case xEntryType
    Case EntryType.Register
        Set MakeDataEntryQueries = MakeInsertQueries(xEntryData)
    Case EntryType.Update
        Set MakeDataEntryQueries = MakeUpdateQueries(xEntryData)
    Case EntryType.Delete
        Set MakeDataEntryQueries = MakeDeleteQueries(xEntryData)
    End Select
End Function


'====================================================================================================
' Insertクエリを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Private Function MakeInsertQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstInsertQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${columns}", GetColumnPhrase(xEntryData.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeInsertQueries = queries
End Function


'====================================================================================================
' Updateクエリを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Private Function MakeUpdateQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstUpdateQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${set}", GetSetPhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeUpdateQueries = queries
End Function


'====================================================================================================
' Deleteクエリを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Private Function MakeDeleteQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstDeleteQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
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
