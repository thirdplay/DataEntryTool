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
' �e�[�u�����
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mTableName As String                ' �e�[�u����
Private mColumnInfoList As Collection       ' �J������񃊃X�g


'====================================================================================================
' �e�[�u�����̎擾/�ݒ�
'====================================================================================================
Public Property Get TableName() As String
    TableName = mTableName
End Property
Public Property Let TableName(TableName As String)
    mTableName = TableName
End Property


'====================================================================================================
' �J������񃊃X�g�̎擾/�ݒ�
'====================================================================================================
Public Property Get ColumnInfoList() As Collection
    Set ColumnInfoList = mColumnInfoList
End Property
Public Property Let ColumnInfoList(ColumnInfoList As Collection)
    Set mColumnInfoList = ColumnInfoList
End Property
