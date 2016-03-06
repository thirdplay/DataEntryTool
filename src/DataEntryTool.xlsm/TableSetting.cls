VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' テーブル設定
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mRow AS String                      ' 行番号
Private mPhysicsName AS String              ' 物理名
Private mLogicalName As String              ' 論理名
Private mDataEntryTarget As String          ' データ投入対象
Private mProcessCount As String             ' 処理件数


'====================================================================================================
' 行番号の取得/設定
'====================================================================================================
Public Property Get Row() As String
    Row = mRow
End Property
Public Property Let Row(Row As String)
    mRow = Row
End Property


'====================================================================================================
' 物理名の取得/設定
'====================================================================================================
Public Property Get PhysicsName() As String
    PhysicsName = mPhysicsName
End Property
Public Property Let PhysicsName(PhysicsName As String)
    mPhysicsName = PhysicsName
End Property


'====================================================================================================
' 論理名の取得/設定
'====================================================================================================
Public Property Get LogicalName() As String
    LogicalName = mLogicalName
End Property
Public Property Let LogicalName(LogicalName As String)
    mLogicalName = LogicalName
End Property


'====================================================================================================
' データ投入対象の取得/設定
'====================================================================================================
Public Property Get DataEntryTarget() As String
    DataEntryTarget = mDataEntryTarget
End Property
Public Property Let DataEntryTarget(DataEntryTarget As String)
    mDataEntryTarget = DataEntryTarget
End Property


'====================================================================================================
' 処理件数の取得/設定
'====================================================================================================
Public Property Get ProcessCount() As String
    ProcessCount = mProcessCount
End Property
Public Property Let ProcessCount(ProcessCount As String)
    mProcessCount = ProcessCount
End Property
