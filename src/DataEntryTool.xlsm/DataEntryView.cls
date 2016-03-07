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
' 名前定義
Private Const cstDatabaseType = "DatabaseType"          ' データベース種類
Private Const cstServerName = "ServerName"              ' サーバ名
Private Const cstPort = "Port"                          ' ポート
Private Const cstDatabaseName = "DatabaseName"          ' データベース名
Private Const cstUserId = "UserId"                      ' ユーザID
Private Const cstPassword = "Password"                  ' パスワード
Private Const cstTableBase = "TableBase"                ' テーブル一覧の基準セル
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
    Comments = 0            ' コメント
    ColumnName              ' 列名
    DataType                ' データ型
    DataLength              ' データ長
    IsRequired              ' 必須指定
    IsPrimaryKey            ' 主キー指定
    Max
End Enum


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
' セル範囲の変更イベント
'----------------------------------------------------------------------------------------------------
' IN : sheet ワークシート
'    : target 対象のセル範囲
'====================================================================================================
Public Sub ChangeCellRange(sheet As Object, ByVal target As Range)
    With sheet
        If Not Intersect(target, .Range(cstDatabaseType)) Is Nothing Then
            ' PostgreSQL
            If target.Value = cstDatabaseTypePostgreSQL Then
                .Range(cstPort).Rows.Hidden = False
                .Range(cstDatabaseName).Rows.Hidden = False
            ' PostgreSQL以外
            Else
                .Range(cstPort).Rows.Hidden = True
                .Range(cstDatabaseName).Rows.Hidden = True
            End If
        End If
    End With
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
        End With

        ' 入力チェック
        Call CheckInputValue(mSetting.DatabaseType, "データベース種類")
        If mSetting.DatabaseType = cstDatabaseTypeOracle Then
            Call CheckInputValue(mSetting.ServerName, "サーバ名")
            Call CheckInputValue(mSetting.UserId, "ユーザID")
            Call CheckInputValue(mSetting.Password, "パスワード")
        ElseIf mSetting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Call CheckInputValue(mSetting.ServerName, "サーバ名")
            Call CheckInputValue(mSetting.Port, "ポート")
            Call CheckInputValue(mSetting.DatabaseName, "データベース名")
            Call CheckInputValue(mSetting.UserId, "ユーザID")
            Call CheckInputValue(mSetting.Password, "パスワード")
        End If
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
' IN : criteriaFlg 取得条件フラグ(true:対象のみ,false:全て)
' OUT: テーブル設定リスト
'====================================================================================================
Public Function GetTableSettingsByCriteria(criteriaFlg As Boolean) As Collection
    Dim rowIndex As Long
    Dim ts As TableSetting
    Dim isTarget As Boolean
    Dim list As Collection
    Set list = New Collection

    With ThisWorkbook.Worksheets(cstSheetMain)
        rowIndex = .Range(cstTableBase).Row + 1
        Do While .Cells(rowIndex, TableSettingCol.PhysicsName).Value <> ""
            ' 取得条件フラグがtrueかつ、データ投入対象列が空の場合、対象外
            isTarget = True
            If criteriaFlg And .Cells(rowIndex, TableSettingCol.DataEntryTarget).Value = "" Then
                isTarget = False
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
            ReDim columnRange(ColumnDefinitionRow.Max, td.ColumnDefinitions.Count)
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
'====================================================================================================
'GetTableDataList


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

