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
' �e�[�u����`
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mTableName AS String                    ' �e�[�u����
Private mColumnDefinitions As Collection        ' �J������`���X�g


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
' �J������`���X�g�̎擾/�ݒ�
'====================================================================================================
Public Property Get ColumnDefinitions() As Collection
    Set ColumnDefinitions = mColumnDefinitions
End Property
Public Property Let ColumnDefinitions(ColumnDefinitions As Collection)
    Set mColumnDefinitions = ColumnDefinitions
End Property