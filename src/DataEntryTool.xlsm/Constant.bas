Attribute VB_Name = "Constant"
Option Explicit
Option Private Module

'====================================================================================================
'
' �萔���W���[��
'
'====================================================================================================
' �f�[�^�x�[�X���
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' ���s�R�[�h
Public Const cstLinefeedCodeCRLF = "CRLF"                   ' CRLF
Public Const cstLinefeedCodeLF = "LF"                       ' LF

' �������
Public Enum EntryType
    Register = 0        ' �o�^
    Update              ' �X�V
    Remove              ' �폜
End Enum
