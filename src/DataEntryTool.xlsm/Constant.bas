Attribute VB_Name = "Constant"
Option Explicit
Option Private Module

'====================================================================================================
'
' �萔���W���[��
'
'====================================================================================================
' �V�[�g��
Public Const cstSheetMain = "�f�[�^�����c�[��"              ' ���C���V�[�g
Public Const cstSheetTemplate = "�e���v���[�g"              ' �e���v���[�g�V�[�g

' ���O��`
Public Const cstTableBase = "TableBase"                     ' �e�[�u���ꗗ�̊�Z��

' �f�[�^�x�[�X���
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' ���s�R�[�h
Public Const cstLinefeedCodeCRLF = "CRLF"                   ' CRLF
Public Const cstLinefeedCodeLF = "LF"                       ' LF

' �}�N�����
Public Enum MacroType
    Database = &H01                     ' �f�[�^�x�[�X
    DataEntry = &H02 Or Database        ' �f�[�^����
    Setting = &H04                      ' �ݒ�
End Enum

' �e�[�u���ݒ�̗�Q�ƒl
Public Enum TableSettingCol
    PhysicsName = 1         ' ������
    LogicalName             ' �_����
    DataEntryTarget         ' �f�[�^����Ώ�
    ProcessCount            ' ��������
    Max = ProcessCount
End Enum

' �J������`�̍s�Q�ƒl
Public Enum ColumnDefinitionRow
    Comments = 1            ' �R�����g
    ColumnName              ' ��
    DataType                ' �f�[�^�^
    DataLength              ' �f�[�^��
    IsRequired              ' �K�{�w��
    IsPrimaryKey            ' ��L�[�w��
    Max = IsPrimaryKey
End Enum
' �e�[�u�����R�[�h�̊J�n�s
Public Const cstTableRecordBase = ColumnDefinitionRow.Max + 1

' �J�[�\�����P�[�V����
Public Const adUseClient = 3
