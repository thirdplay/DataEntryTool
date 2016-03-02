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
' �J�������
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mOrdinalPosition As Long        ' ��
Private mColumnName As String           ' ���ږ�
Private mComment As String              ' �R�����g
Private mDataType As String             ' �f�[�^�^
Private mDataLength As String           ' �f�[�^��
Private mIsNullable As String           ' NULL�w��\(0:NULL�s��,1:NULL�\)
Private mIsPrimaryKey As String         ' ��L�[�w��(0:��L�[�ȊO,1:��L�[)


'====================================================================================================
' �񏇂̎擾/�ݒ�
'====================================================================================================
Public Property Get OrdinalPosition() As Long
    OrdinalPosition = mOrdinalPosition
End Property
Public Property Let OrdinalPosition(OrdinalPosition As Long)
    mOrdinalPosition = OrdinalPosition
End Property


'====================================================================================================
' ���ږ��̎擾/�ݒ�
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
Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(Comment As String)
    mComment = Comment
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
