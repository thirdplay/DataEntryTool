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
    PhysicsName = 1
    LogicName
    Target
    Result
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
' OUT:テーブル名一覧
'====================================================================================================
Public Function GetTableNameList() As Collection
    Dim list As Collection
    Dim index As Long
    
    With ThisWorkbook.Worksheets(cstSheetMain)
        Set list = New Collection
        index = .Range(cstTableBase).Row + 1
        
        ' 対象列に値が設定されているテーブル名を追加する
        Do While .Cells(index, TableListIndex.PhysicsName).Value <> ""
            Call list.Add(.Cells(index, TableListIndex.PhysicsName).Value)
            index = index + 1
        Loop
    End With
    Set GetTableNameList = list
    
End Function


'====================================================================================================
' テーブル定義シートの再作成
'----------------------------------------------------------------------------------------------------
' IN : tableInfoList テーブル情報リスト
'====================================================================================================
Public Sub RecreateTableDefinition(tableInfoList As Collection)
On Error GoTo Finally
    Dim ti As TableInfo
    Dim ci As ColumnInfo
    Dim tmplSheet As Object
    Set tmplSheet = ThisWorkbook.Worksheets(cstSheetTemplate)
    tmplSheet.Visible = True

    For Each ti In TableInfoList
        ' 既にシートが存在する場合は削除する
        If ExistsSheet(ti.TableName) Then
            ThisWorkbook.Worksheets(ti.TableName).Delete
        End If

        ' テンプレートシートをコピーする
        tmplSheet.Copy Before := tmplSheet
        ThisWorkbook.ActiveSheet.Name = ti.TableName

        ' コピーしたシートにカラム情報を書き込む
        With ThisWorkbook.ActiveSheet
            For Each ci In ti.ColumnInfoList
                .Cells(1, ci.OrdinalPosition).Value = ci.Comment
                .Cells(2, ci.OrdinalPosition).Value = ci.ColumnName
                .Cells(3, ci.OrdinalPosition).Value = ci.DataType
                .Cells(4, ci.OrdinalPosition).Value = ci.DataLength
                .Cells(5, ci.OrdinalPosition).Value = ci.IsNullable
                .Cells(6, ci.OrdinalPosition).Value = ci.IsPrimaryKey
            Next
        End With
    Next
Finally:
    ThisWorkbook.Worksheets(cstSheetMain).Activate
    tmplSheet.Visible = False
    Set ti = Nothing
    Set ci = Nothing
    Set tmplSheet = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


' テーブルデータ一覧取得
'GetTableDataList


'====================================================================================================
' 入力値をチェックします。
'----------------------------------------------------------------------------------------------------
' IN : value 入力値
'    : itemName 項目名
'====================================================================================================
Private Sub CheckInputValue(value As String, itemName As String)
    If value = "" Then
        Err.Raise 100, , itemName & "を入力してください。"
    End If
End Sub


'====================================================================================================
' シートが存在するか判定します。
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

