VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' カラム定義
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mColumnId As Long               ' 項目ID
Private mColumnName As String           ' 列名
Private mComments As String             ' コメント
Private mDataType As String             ' データ型
Private mDataLength As String           ' データ長
Private mIsRequired As String           ' 必須指定(0:任意,1:必須)
Private mIsPrimaryKey As String         ' 主キー指定(0:主キー以外,1:主キー)


'====================================================================================================
' 列IDの取得/設定
'====================================================================================================
Public Property Get ColumnId() As Long
    ColumnId = mColumnId
End Property
Public Property Let ColumnId(ColumnId As Long)
    mColumnId = ColumnId
End Property


'====================================================================================================
' 列名の取得/設定
'====================================================================================================
Public Property Get ColumnName() As String
    ColumnName = mColumnName
End Property
Public Property Let ColumnName(ColumnName As String)
    mColumnName = ColumnName
End Property


'====================================================================================================
' コメントの取得/設定
'====================================================================================================
Public Property Get Comments() As String
    Comments = mComments
End Property
Public Property Let Comments(Comments As String)
    mComments = Comments
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
Public Property Get IsRequired() As String
    IsRequired = mIsRequired
End Property
Public Property Let IsRequired(IsRequired As String)
    mIsRequired = IsRequired
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
