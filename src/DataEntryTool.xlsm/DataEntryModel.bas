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
' 投入データを取得します
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'    : xTableName テーブル名
' OUT: 投入データ
'====================================================================================================
Public Function GetEntryData(xEntryType As EntryType, xTableName As String) As EntryData
    Dim ed As EntryData
    Dim cd As ColumnDefinition
    Dim rowIndex As Long
    Dim dataCount As Long
    Dim columnRange As Range
    Dim cr As Range

    With ThisWorkbook.Worksheets(xTableName)
        Set ed = New EntryData
        ed.EntryType = xEntryType
        ed.TableName = xTableName

        ' カラム定義リストの作成
        ed.ColumnDefinitions = New Collection
        dataCount = WorkSheetEx.GetColDataCount(xTableName, ColumnDefinitionRow.ColumnName, 1)
        Set columnRange = .Range(.Cells(1, 1), .Cells(ColumnDefinitionRow.Max, dataCount))
        For Each cr In columnRange.Columns
            Set cd = New ColumnDefinition
            cd.ColumnName = cr.Cells(ColumnDefinitionRow.ColumnName, 1)
            cd.DataType = cr.Cells(ColumnDefinitionRow.DataType, 1)
            cd.IsPrimaryKey = cr.Cells(ColumnDefinitionRow.IsPrimaryKey, 1)
            Call ed.ColumnDefinitions.Add(cd)
        Next

        ' レコード範囲の設定
        dataCount = WorkSheetEx.GetRowDataCount(xTableName, cstTableRecordBase, 1)
        ed.RecordRange = .Range(.Cells(cstTableRecordBase, 1), .Cells(cstTableRecordBase + dataCount - 1, ed.ColumnDefinitions.Count))

        Set GetEntryData = ed
    End With
End Function


'====================================================================================================
' データ投入を実行します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: 処理件数
'====================================================================================================
Public Function ExecuteDataEntry(xEntryData As EntryData) As Long
On Error GoTo ErrHandler
    Dim i As Long
    Dim query As String
    Dim procCnt As Long

    ' トランザクション開始
    Call Database.BeginTrans

    ' クエリ生成
    query = MakeQuery(xEntryData)

    ' クエリ実行
    procCnt = Database.ExecuteQuery(query)

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
        "データ行:" & (cstTableRecordBase + i - 1) & vbNewLine & _
        "[エラー詳細]" & vbNewLine & _
        Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function


'====================================================================================================
' クエリを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリ文字列
'====================================================================================================
Private Function MakeQuery(xEntryData As EntryData) As String
    Select Case xEntryData.EntryType
    Case EntryType.Register
        MakeQuery = MakeInsertQuery(xEntryData)
    Case EntryType.Update
        MakeQuery = MakeUpdateQuery(xEntryData)
    Case EntryType.Remove
        MakeQuery = MakeDeleteQuery(xEntryData)
    End Select
End Function


'====================================================================================================
' Insertクエリリストを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリ文字列
'====================================================================================================
Private Function MakeInsertQuery(xEntryData As EntryData) As String
    Dim rr As Range
    Dim query As String

    For Each rr In xEntryData.RecordRange.Rows
        query = query & cstInsertQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${columns}", GetColumnPhrase(xEntryData.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(xEntryData.ColumnDefinitions, rr))
        query = query & Database.GetQuerySuffix()
    Next
    MakeInsertQuery = query
End Function


'====================================================================================================
' Updateクエリリストを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリ文字列
'====================================================================================================
Private Function MakeUpdateQuery(xEntryData As EntryData) As String
    Dim rr As Range
    Dim query As String

    For Each rr In xEntryData.RecordRange.Rows
        query = query & cstUpdateQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${set}", GetSetPhrase(xEntryData.ColumnDefinitions, rr))
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, rr))
        query = query & Database.GetQuerySuffix()
    Next
    MakeUpdateQuery = query
End Function


'====================================================================================================
' Deleteクエリリストを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Private Function MakeDeleteQuery(xEntryData As EntryData) As String
    Dim rr As Range
    Dim query As String

    For Each rr In xEntryData.RecordRange.Rows
        query = query & cstDeleteQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, rr))
        query = query & Database.GetQuerySuffix()
    Next
    MakeDeleteQuery = query
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
        result = result & GetItemValue(record.Cells(1, i), cd(i).DataType) & cstDelimiter
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
            result = Replace(result, "${value}", GetItemValue(record.Cells(1, i), cd(i).DataType))
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
            result = Replace(result, "${value}", GetItemValue(record.Cells(1, i), cd(i).DataType))
            result = result & cstWherePhraseDelimiter
        End If
    Next
    GetWherePhrase = Left(result, Len(result) - Len(cstWherePhraseDelimiter))
End Function


'====================================================================================================
' データ型に対応したデータ値の項目値を取得します
'----------------------------------------------------------------------------------------------------
' IN : dataValue データ値
'    : xDataType データ型
' OUT: Value句
'====================================================================================================
Private Function GetItemValue(ByVal dataValue As String, ByVal xDataType As String) As String
    Dim itemValue As String

    itemValue = dataValue
    If itemValue = "" Then
        itemValue = "NULL"
    ' 文字列
    ElseIf Database.IsDataTypeString(xDataType) Then
        itemValue = Replace(itemValue, "'", "''")                                   ' 単一引用符エスケープ
        itemValue = Replace(itemValue, vbLf, "'" & Setting.LinefeedCode & "'")      ' 改行コード変換
        itemValue = "'" & itemValue & "'"                                           ' 単一引用符付与
    ' 日付
    ElseIf Database.IsDataTypeDate(xDataType) Then
        itemValue = "TO_DATE('" & itemValue & "','" & Setting.DateFormat & "')"
    ' タイムスタンプ
    ElseIf Database.IsDataTypeTimestamp(xDataType) Then
        itemValue = "TO_TIMESTAMP('" & itemValue & "','" & Setting.TimestampFormat & "')"
    End If
    GetItemValue = itemValue
End Function
