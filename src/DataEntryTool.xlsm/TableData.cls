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
' �����f�[�^
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mColumnDefinitions As Collection    ' �J������`���X�g
Private mRecordRange As Range               ' ���R�[�h�͈�


'====================================================================================================
' �J������`���X�g�̎擾/�ݒ�
'====================================================================================================
Public Property Get ColumnDefinitions() As Collection
    Set ColumnDefinitions = mColumnDefinitions
End Property
Public Property Let ColumnDefinitions(ColumnDefinitions As Collection)
    Set mColumnDefinitions = ColumnDefinitions
End Property


'====================================================================================================
' ���R�[�h�͈͂̎擾/�ݒ�
'====================================================================================================
Public Property Get RecordRange() As Range
    Set RecordRange = mRecordRange
End Property
Public Property Let RecordRange(RecordRange As Range)
    Set mRecordRange = RecordRange
End Property
