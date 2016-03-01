VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' カラム情報
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mLogicalName As String      ' 論理名
Private mPhysicsName As String      ' 物理名
Private mDataType As String         ' データ型
Private mDataLength As String       ' データ長
Private mIsNullable As String       ' NULL指定可能(0:NULL不可,1:NULL可能)
Private mIsPrimaryKey As String     ' 主キー指定(0:主キー以外,1:主キー)


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
' 物理名の取得/設定
'====================================================================================================
Public Property Get PhysicsName() As String
    PhysicsName = mPhysicsName
End Property
Public Property Let PhysicsName(PhysicsName As String)
    mPhysicsName = PhysicsName
End Property


'====================================================================================================
' データ型の取得/設定
'====================================================================================================
Public Property Get DataType() As String
    DataType = mDataType
End Property
Public Property Let DataType(DataType As String)
    mDataType = DataType
End Property


'====================================================================================================
' データ長の取得/設定
'====================================================================================================
Public Property Get DataLength() As String
    DataLength = mDataLength
End Property
Public Property Let DataLength(DataLength As String)
    mDataLength = DataLength
End Property


'====================================================================================================
' NULL指定可能の取得/設定
'====================================================================================================
Public Property Get IsNullable() As String
    IsNullable = mIsNullable
End Property
Public Property Let IsNullable(IsNullable As String)
    mIsNullable = IsNullable
End Property


'====================================================================================================
' 主キー指定の取得/設定
'====================================================================================================
Public Property Get IsPrimaryKey() As String
    IsPrimaryKey = mIsPrimaryKey
End Property
Public Property Let IsPrimaryKey(IsPrimaryKey As String)
    mIsPrimaryKey = IsPrimaryKey
End Property

