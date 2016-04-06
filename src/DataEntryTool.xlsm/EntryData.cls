VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntryData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' 投入データ
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mEntryType As EntryType             ' 投入種類
Private mTableName As String                ' テーブル名
Private mColumnDefinitions As Collection    ' カラム定義リスト
Private mRecordRange As Range               ' レコード範囲


'====================================================================================================
' 投入種類の取得/設定
'====================================================================================================
Public Property Get EntryType() As EntryType
    EntryType = mEntryType
End Property
Public Property Let EntryType(xEntryType As EntryType)
    mEntryType = xEntryType
End Property


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


'====================================================================================================
' レコード範囲の取得/設定
'====================================================================================================
Public Property Get RecordRange() As Range
    Set RecordRange = mRecordRange
End Property
Public Property Let RecordRange(RecordRange As Range)
    Set mRecordRange = RecordRange
End Property

