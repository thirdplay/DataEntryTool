Attribute VB_Name = "Constant"
Option Explicit
Option Private Module

'====================================================================================================
'
' �萔���W���[��
'
'====================================================================================================
' ���O��`
Public Const cstDatabaseType = "DatabaseType"               ' �f�[�^�x�[�X���
Public Const cstServerName = "ServerName"                   ' �T�[�o��
Public Const cstPort = "Port"                               ' �|�[�g
Public Const cstDatabaseName = "DatabaseName"               ' �f�[�^�x�[�X��
Public Const cstUserId = "UserId"                           ' ���[�UID
Public Const cstPassword = "Password"                       ' �p�X���[�h
Public Const cstLinefeedCode = "LinefeedCode"               ' ���s�R�[�h
Public Const cstDateFormat = "DateFormat"                   ' ���t����
Public Const cstTableBase = "TableBase"                     ' �e�[�u���ꗗ�̊�Z��

' �f�[�^�x�[�X���
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' �������
Public Enum EntryType
    Register = 0        ' �o�^
    Update              ' �X�V
    Remove              ' �폜
End Enum
