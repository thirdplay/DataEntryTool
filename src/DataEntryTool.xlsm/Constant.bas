Attribute VB_Name = "Constant"
Option Explicit

'====================================================================================================
'
' �萔���W���[��
'
'====================================================================================================
' �f�[�^�x�[�X���
Public Const cstDatabaseTypeOracle = "Oracle"               ' Oracle
Public Const cstDatabaseTypePostgreSQL = "PostgreSQL"       ' PostgreSQL

' �������
Public Enum EntryType
    Register = 0        ' �o�^
    Update              ' �X�V
    Remove              ' �폜
End Enum
