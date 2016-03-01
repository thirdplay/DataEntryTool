VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OracleDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' OracleDao
'
'====================================================================================================
Implements IDataEntryDao


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Dim mCononect As Object     ' �ڑ��I�u�W�F�N�g


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Private Sub Class_Initialize()
    Const RefFile As String = "C:\Program Files\Common Files\System\ado\msado15.dll"
    ActiveWorkbook.VBProject.References.AddFromFile RefFile
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Private Sub Class_Terminate()
    Call IDataEntryDao_Disconnect()
End Sub


'====================================================================================================
' DB�ڑ�
'----------------------------------------------------------------------------------------------------
' IN : databaseInfo �f�[�^�x�[�X���
'====================================================================================================
Public Sub IDataEntryDao_Connect(databaseInfo As DatabaseInfo)
On Error GoTo ErrHandler
    Set mCononect = New ADODB.Connection

    mCononect.Open "Driver={Microsoft ODBC for Oracle};" _
        & "Server=" & databaseInfo.ServerName & ";" _
        & "UID=" & databaseInfo.UserId & ";" _
        & "PWD=" & databaseInfo.Password & ";"
    Exit Sub
ErrHandler:
    MsgBox "DB�ڑ����s" & vbNewLine & Err.Description
    Set mCononect = Nothing
End Sub


'====================================================================================================
' DB�ؒf
'====================================================================================================
Public Sub IDataEntryDao_Disconnect()
    If Not mCononect Is Nothing Then
        mCononect.Close
        Set mCononect = Nothing
    End If
End Sub


'====================================================================================================
' �g�����U�N�V�����J�n
'====================================================================================================
Public Sub IDataEntryDao_BeginTrans()
    MsgBox "Oracle:BeginTrans"
End Sub


'====================================================================================================
' �R�~�b�g
'====================================================================================================
Public Sub IDataEntryDao_CommitTrans()
    MsgBox "Oracle:CommitTrans"
End Sub


'====================================================================================================
' ���[���o�b�N
'====================================================================================================
Public Sub IDataEntryDao_RollbackTrans()
    MsgBox "Oracle:RollbackTrans"
End Sub


'====================================================================================================
' �w�肳�ꂽ�e�[�u���̃J������`���擾���܂��B
'----------------------------------------------------------------------------------------------------
' IN : tableName �e�[�u����
' OUT: �J������񃊃X�g
'====================================================================================================
'Public Function GetTableInfo(tableName As String)
'select
'UCC.COMMENTS as LOGICAL_NAME
',UTC.COLUMN_NAME as PHYSICS_NAME
',UTC.DATA_TYPE
',case UTC.DATA_TYPE
'  when 'CHAR' then TO_CHAR(UTC.DATA_LENGTH)
'  when 'VARCHAR2' then TO_CHAR(UTC.DATA_LENGTH)
'  when 'NUMBER' then UTC.DATA_PRECISION || ',' || UTC.DATA_SCALE
'  else ''
'end as DATA_LENGTH
',DECODE(UTC.NULLABLE,'N','0','1') as IS_NULLABLE
',DECODE(PKC.TABLE_NAME, NULL, '0', '1') as IS_PRIMARY_KEY
'from USER_TAB_COLUMNS UTC
'  left join USER_COL_COMMENTS UCC on(
'    UCC.TABLE_NAME = UTC.TABLE_NAME
'    and UCC.COLUMN_NAME = UTC.COLUMN_NAME
'  )
'  left join (
'    select
'      CC.TABLE_NAME
'      ,CC.COLUMN_NAME
'    from USER_CONS_COLUMNS CC
'      inner join USER_CONSTRAINTS CONS on(
'        CONS.CONSTRAINT_NAME = CC.CONSTRAINT_NAME
'        and CONS.CONSTRAINT_TYPE = 'P'
'      )
'  ) PKC on (
'    PKC.TABLE_NAME = UTC.TABLE_NAME
'    and PKC.COLUMN_NAME = UTC.COLUMN_NAME
'  )
'where UTC.TABLE_NAME='RCAT_M_USER'
'order by UTC.COLUMN_ID
'
'    Set GetTableInfo = Nothing
'End Sub


'====================================================================================================
' �N�G�����s
'----------------------------------------------------------------------------------------------------
' IN : query �N�G��������
'====================================================================================================
'Public Sub Execute(query)
'End Sub