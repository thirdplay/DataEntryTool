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
Public Const cstDatabaseType = "DatabaseType"               ' �f�[�^�x�[�X���
Public Const cstServerName = "ServerName"                   ' �T�[�o��
Public Const cstServerNameLabel = "ServerNameLabel"         ' �T�[�o�����x��
Public Const cstUserId = "UserId"                           ' ���[�UID
Public Const cstPassword = "Password"                       ' �p�X���[�h
Public Const cstPort = "Port"                               ' �|�[�g
Public Const cstDatabaseName = "DatabaseName"               ' �f�[�^�x�[�X��
Public Const cstLinefeedCode = "LinefeedCode"               ' ���s�R�[�h
Public Const cstDateFormat = "DateFormat"                   ' ���t����
Public Const cstTimestampFormat = "TimestampFormat"         ' �^�C���X�^���v����
Public Const cstOutputDir = "OutputDir"                     ' �o�͐�
Public Const cstTableBase = "TableBase"                     ' �e�[�u���ꗗ�̊�Z��

' �f�[�^�x�[�X���
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' �T�[�o�����x��
Public Const cstServerNameLabelOracle = "�ڑ����ʎq"
Public Const cstServerNameLabelPostgreSQL = "�z�X�g��"

' ���s�R�[�h
Public Const cstLinefeedCodeCRLF = "CRLF"                   ' CRLF
Public Const cstLinefeedCodeLF = "LF"                       ' LF

' �G���[�ԍ�
Public Enum ErrNumber
    Error = 1000        ' ���[�U��`�̃G���[
    Warning = 2000      ' ���[�U��`�̌x��
    Abort = 3000        ' ���[�U��`�̒��f
End Enum

' �ݒ���
Public Enum SettingType
    None = &H0                          ' �Ȃ�
    Database = &H1                      ' �f�[�^�x�[�X
    DataEntry = &H2 Or Database         ' �f�[�^����
End Enum

' �������
Public Enum EntryType
    Register = 0        ' �o�^
    Update              ' �X�V
    Delete              ' �폜
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
' �R�}���h��񓯊��Ɏ��s���邱�Ƃ�����
Public Const adAsyncExecute = &H10
