Attribute VB_Name = "DataEntryModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' データ投入モデルのモジュール
'
'====================================================================================================

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

    ' クエリ生成
    Select Case xEntryType
    Case EntryType.Register
        Set queries= xEntryData.MakeInsertQueries()
    Case EntryType.Update
        Set queries = xEntryData.MakeUpdateQueries()
    Case EntryType.Delete
        Set queries = xEntryData.MakeDeleteQueries()
    End Select

    ' データ投入
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
