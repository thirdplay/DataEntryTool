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
' �J������`
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mColumnId As Long               ' ����ID
Private mColumnName As String           ' ��
Private mComments As String             ' �R�����g
Private mDataType As String             ' �f�[�^�^
Private mDataLength As String           ' �f�[�^��
Private mIsNullable As String           ' NULL�w��\(0:NULL�s��,1:NULL�\)
Private mIsPrimaryKey As String         ' ��L�[�w��(0:��L�[�ȊO,1:��L�[)


'====================================================================================================
' ��ID�̎擾/�ݒ�
'====================================================================================================
Public Property Get ColumnId() As Long
    ColumnId = mColumnId    
End Property
Public Property Let ColumnId(ColumnId As Long)
    mColumnId = ColumnId
End Property


'====================================================================================================
' �񖼂̎擾/�ݒ�
'====================================================================================================
Public Property Get ColumnName() As String
    ColumnName = mColumnName
End Property
Public Property Let ColumnName(ColumnName As String)
    mColumnName = ColumnName
End Property


'====================================================================================================
' �R�����g�̎擾/�ݒ�
'====================================================================================================
Public Property Get Comments() As String
    Comments = mComments
End Property
Public Property Let Comments(Comments As String)
    mComments = Comments
End Property


'====================================================================================================
' �f�[�^�^�̎擾/�ݒ�
'====================================================================================================
Public Property Get DataType() As String
    DataType = mDataType
End Property
Public Property Let DataType(DataType As String)
    mDataType = DataType
End Property


'====================================================================================================
' �f�[�^���̎擾/�ݒ�
'====================================================================================================
Public Property Get DataLength() As String
    DataLength = mDataLength
End Property
Public Property Let DataLength(DataLength As String)
    mDataLength = DataLength
End Property


'====================================================================================================
' NULL�w��\�̎擾/�ݒ�
'====================================================================================================
Public Property Get IsNullable() As String
    IsNullable = mIsNullable
End Property
Public Property Let IsNullable(IsNullable As String)
    mIsNullable = IsNullable
End Property


'====================================================================================================
' ��L�[�w��̎擾/�ݒ�
'====================================================================================================
Public Property Get IsPrimaryKey() As String
    IsPrimaryKey = mIsPrimaryKey
End Property
Public Property Let IsPrimaryKey(IsPrimaryKey As String)
    mIsPrimaryKey = IsPrimaryKey
End Property