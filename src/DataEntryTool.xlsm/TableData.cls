VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableData"
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
Private mColumnDefinitions As Collection    ' カラム定義リスト
Private mRecordRange As Range               ' レコード範囲


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

