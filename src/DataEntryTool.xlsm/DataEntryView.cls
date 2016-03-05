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

' テーブル一覧の参照値
Private Enum TableListIndex
    PhysicsName = 1         ' 物理名
    LogicName               ' 論理名
    Target                  ' 対象
    ProcessCount            ' 処理件数
    Max = ProcessCount
End Enum


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mSetting As Setting     ' 設定


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
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
    'Call CheckInputValue(mSetting.DatabaseType, "データベース種類")
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
' OUT:設定
'====================================================================================================
Public Function GetSetting() As Setting
    Set GetSetting = mSetting
End Function


'====================================================================================================
' テーブル名一覧を返却します
'----------------------------------------------------------------------------------------------------
' OUT:テーブル情報一覧
'====================================================================================================
Public Function GetTableNameList() As Collection
    Dim list As Collection
    Dim rowIndex As Long
    
    With ThisWorkbook.Worksheets(cstSheetMain)
        Set list = New Collection
        rowIndex = .Range(cstTableBase).Row + 1

        Do While .Cells(rowIndex, TableListIndex.PhysicsName).Value <> ""
            Call list.Add(.Cells(rowIndex, TableListIndex.PhysicsName).Value)
            rowIndex = rowIndex + 1
        Loop
    End With
    Set GetTableNameList = list
    
End Function


'====================================================================================================
' テーブルシートの作成
'----------------------------------------------------------------------------------------------------
' IN : tableDefinitions テーブル情報リスト
'====================================================================================================
Public Sub CreateTableSheet(tableDefinitions As Collection)
On Error GoTo Finally
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim columnRange As Variant
    Dim tmplSheet As Object
    Set tmplSheet = ThisWorkbook.Worksheets(cstSheetTemplate)
    tmplSheet.Visible = True

    For Each td In tableDefinitions
        ' 既にシートが存在する場合は削除する
        If ExistsSheet(td.TableName) Then
            ThisWorkbook.Worksheets(td.TableName).Delete
        End If

        ' テンプレートシートをコピーする
        tmplSheet.Copy Before := tmplSheet
        ThisWorkbook.ActiveSheet.Name = td.TableName

        ' コピーしたシートにカラム情報を書き込む
        If td.ColumnDefinitions.Count > 0 Then
            ReDim columnRange(6, td.ColumnDefinitions.Count)
            For Each cd In td.ColumnDefinitions
                columnRange(0, cd.ColumnId) = cd.Comments
                columnRange(1, cd.ColumnId) = cd.ColumnName
                columnRange(2, cd.ColumnId) = cd.DataType
                columnRange(3, cd.ColumnId) = cd.DataLength
                columnRange(4, cd.ColumnId) = cd.IsNullable
                columnRange(5, cd.ColumnId) = cd.IsPrimaryKey
            Next
            ThisWorkbook.ActiveSheet.Range(Cells(1, 1), Cells(6, td.ColumnDefinitions.Count)) = columnRange
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

