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
Private mDatabaseInfo As DatabaseInfo           ' データベース情報


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
    Set mDatabaseInfo = New DatabaseInfo
    With mDatabaseInfo
        .ServerName = ThisWorkbook.Worksheets(cstSheetMain).Range(cstServerName).Value
        .Port = ThisWorkbook.Worksheets(cstSheetMain).Range(cstPort).Value
        .DatabaseName = ThisWorkbook.Worksheets(cstSheetMain).Range(cstDatabaseName).Value
        .UserId = ThisWorkbook.Worksheets(cstSheetMain).Range(cstUserId).Value
        .Password = ThisWorkbook.Worksheets(cstSheetMain).Range(cstPassword).Value
    End With
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Set mDatabaseInfo = Nothing
End Sub


'====================================================================================================
' データベース情報を返却します
'----------------------------------------------------------------------------------------------------
' OUT:データベース情報
'====================================================================================================
Public Function GetDatabaseInfo() As DatabaseInfo
    Set GetDatabaseInfo = mDatabaseInfo
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
            If .Cells(index, TableListIndex.Target).Value <> "" Then
                list.Add (.Cells(index, TableListIndex.PhysicsName))
            End If
            index = index + 1
        Loop
    End With
    Set GetTableNameList = list
    
End Function

'テーブルデータ一覧取得

'テーブルシートの再作成



