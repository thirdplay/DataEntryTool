VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データ投入ビュー
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
' シート名
Private Const cstSheetMain = "データ投入ツール"         ' メインシート
Private Const cstSheetTemplate = "テンプレート"         ' テンプレートシート

' テーブル設定の列参照値
Private Enum TableSettingCol
    PhysicsName = 1         ' 物理名
    LogicalName             ' 論理名
    DataEntryTarget         ' データ操作対象
    ProcessCount            ' 処理件数
    Max = ProcessCount
End Enum

' カラム定義の行参照値
Private Enum ColumnDefinitionRow
    Comments = 1            ' コメント
    ColumnName              ' 列名
    DataType                ' データ型
    DataLength              ' データ長
    IsRequired              ' 必須指定
    IsPrimaryKey            ' 主キー指定
    Max = IsPrimaryKey
End Enum

' テーブルレコードの開始行
Private Const cstTableRecordBase = ColumnDefinitionRow.Max + 1


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mSetting As Setting     ' 設定


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Set mSetting = Nothing
End Sub


'====================================================================================================
' 設定を返却します
'----------------------------------------------------------------------------------------------------
' OUT: 設定
'====================================================================================================
Public Function GetSetting() As Setting
    If mSetting Is Nothing Then
        Set mSetting = New Setting
        With ThisWorkbook.Worksheets(cstSheetMain)
            mSetting.DatabaseType = .Range(cstDatabaseType).Value
            mSetting.ServerName = .Range(cstServerName).Value
            mSetting.Port = .Range(cstPort).Value
            mSetting.DatabaseName = .Range(cstDatabaseName).Value
            mSetting.UserId = .Range(cstUserId).Value
            mSetting.Password = .Range(cstPassword).Value
            mSetting.LinefeedCode = .Range(cstLinefeedCode).Value
            mSetting.DateFormat = .Range(cstDateFormat).Value
        End With

        ' 入力チェック
        Call CheckInputValue(mSetting.DatabaseType, "データベース種類")
        Call CheckInputValue(mSetting.ServerName, "サーバ名")
        If mSetting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Call CheckInputValue(mSetting.Port, "ポート")
            Call CheckInputValue(mSetting.DatabaseName, "データベース名")
        End If
        Call CheckInputValue(mSetting.UserId, "ユーザID")
        Call CheckInputValue(mSetting.Password, "パスワード")
        Call CheckInputValue(mSetting.LinefeedCode, "改行コード")
        Call CheckInputValue(mSetting.DateFormat, "日付書式")
    End If
    Set GetSetting = mSetting
End Function


'====================================================================================================
' テーブル設定リストを返却します
'----------------------------------------------------------------------------------------------------
' OUT: テーブル設定リスト
'====================================================================================================
Public Function GetTableSettings() As Collection
    Set GetTableSettings = GetTableSettingsByCriteria(False)
End Function


'====================================================================================================
' データ投入対象のテーブル設定リストを返却します
'----------------------------------------------------------------------------------------------------
' IN : isEntryTarget 投入対象フラグ(True:対象のみ,False:全て)
' OUT: テーブル設定リスト
'====================================================================================================
Public Function GetTableSettingsByCriteria(isEntryTarget As Boolean) As Collection
    Dim rowIndex As Long
    Dim ts As TableSetting
    Dim isTarget As Boolean
    Dim tableName As String
    Dim list As Collection
    Set list = New Collection

    With ThisWorkbook.Worksheets(cstSheetMain)
        rowIndex = .Range(cstTableBase).Row + 1
        Do While .Cells(rowIndex, TableSettingCol.PhysicsName).Value <> ""
            ' 投入対象フラグがtrueの場合、投入対象外の場合は処理しない
            tableName = .Cells(rowIndex, TableSettingCol.PhysicsName).Value
            isTarget = True
            If isEntryTarget Then
                If .Cells(rowIndex, TableSettingCol.DataEntryTarget).Value = "" Then
                    isTarget = False
                ElseIf Not ExistsSheet(tableName) Then
                    'Err.Raise 100, , "シート[" & tableName & "]が存在しません。"
                    isTarget = False
                ElseIf ThisWorkbook.Worksheets(tableName).Cells(cstTableRecordBase, 1).Value = "" Then
                    isTarget = False
                End If
            End If

            If isTarget Then
                Set ts = New TableSetting
                ts.Row = rowIndex
                ts.PhysicsName = .Cells(rowIndex, TableSettingCol.PhysicsName).Value
                ts.LogicalName = .Cells(rowIndex, TableSettingCol.LogicalName).Value
                ts.DataEntryTarget = .Cells(rowIndex, TableSettingCol.DataEntryTarget).Value
                ts.ProcessCount = .Cells(rowIndex, TableSettingCol.ProcessCount).Value
                Call list.Add(ts)
            End If
            rowIndex = rowIndex + 1
        Loop
    End With
    Set GetTableSettingsByCriteria = list
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

    Dim requiredDic As Object
    Set requiredDic = CreateObject("Scripting.Dictionary")
    Call requiredDic.Add("1", "必須")
    
    Dim primaryKeyDic As Object
    Set primaryKeyDic = CreateObject("Scripting.Dictionary")
    Call primaryKeyDic.Add("1", "PK")

    Dim tmplSheet As Object
    Set tmplSheet = ThisWorkbook.Worksheets(cstSheetTemplate)
    tmplSheet.Visible = True

    For Each td In tableDefinitions
        ' 既にシートが存在する場合は削除する
        If ExistsSheet(td.TableName) Then
            ThisWorkbook.Worksheets(td.TableName).Delete
        End If

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
' テーブルデータ取得
'----------------------------------------------------------------------------------------------------
' IN : tableName テーブル名
'====================================================================================================
Public Function GetTableData(tableName As String) As TableData
    Dim td As TableData
    Dim cd As ColumnDefinition
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim dataCount As Long
    Dim i As Long
    Dim rs As Collection

    With ThisWorkbook.Worksheets(tableName)
        Set td = New TableData

        ' カラム定義リストの作成
        td.ColumnDefinitions = New Collection
        colIndex = 1
        Do While .Cells(ColumnDefinitionRow.ColumnName, colIndex).Value <> ""
            Set cd = New ColumnDefinition
            cd.ColumnName = .Cells(ColumnDefinitionRow.ColumnName, colIndex).Value
            cd.DataType = .Cells(ColumnDefinitionRow.DataType, colIndex).Value
            cd.IsPrimaryKey = .Cells(ColumnDefinitionRow.IsPrimaryKey, colIndex).Value
            Call td.ColumnDefinitions.Add(cd)
            colIndex = colIndex + 1
        Loop

        ' レコード範囲の設定
        dataCount = GetDataCount(tableName, cstTableRecordBase, 1)
        td.RecordRange = .Range(.Cells(cstTableRecordBase, 1), .Cells(cstTableRecordBase + dataCount, td.ColumnDefinitions.Count))

        Set GetTableData = td
    End With
End Function


'====================================================================================================
' 入力値をチェックします
'----------------------------------------------------------------------------------------------------
' IN : inputValue 入力値
'    : itemName 項目名
'====================================================================================================
Private Sub CheckInputValue(inputValue As String, itemName As String)
    If inputValue = "" Then
        Err.Raise 100, , itemName & "を入力してください。"
    End If
End Sub


'====================================================================================================
' シートが存在するか判定します
'----------------------------------------------------------------------------------------------------
' IN : sheetName シート名
' OUT: 存在する場合はtrue、それ以外はfalse
'====================================================================================================
Private Function ExistsSheet(sheetName As String) As Boolean
    Dim sheet As Object
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.Name = sheetName Then
            ExistsSheet = True
            Exit Function
        End If
    Next
    ExistsSheet = False
End Function


'====================================================================================================
' 指定列上のデータがいくつあるか取得します
'----------------------------------------------------------------------------------------------------
' IN : sheetName シート名
'    : rowIndex 行インデックス
'    : colIndex 列インデックス
' OUT: データ数
'====================================================================================================
Private Function GetDataCount(sheetName As String, rowIndex As Long, colIndex As Long) As Long
    Dim startIndex As Long

    startIndex = rowIndex
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex, colIndex).Value <> ""
            rowIndex = rowIndex + 1
        Loop
    End With
    GetDataCount = rowIndex - startIndex - 1
End Function

