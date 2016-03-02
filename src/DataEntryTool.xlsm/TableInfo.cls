VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' テーブル情報
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mTableName As String                ' テーブル名
Private mColumnInfoList As Collection       ' カラム情報リスト


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
' カラム情報リストの取得/設定
'====================================================================================================
Public Property Get ColumnInfoList() As Collection
    Set ColumnInfoList = mColumnInfoList
End Property
Public Property Let ColumnInfoList(ColumnInfoList As Collection)
    Set mColumnInfoList = ColumnInfoList
End Property

