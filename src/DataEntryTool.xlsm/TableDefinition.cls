VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' テーブル定義
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mTableName As String                    ' テーブル名
Private mColumnDefinitions As Collection        ' カラム定義リスト


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
