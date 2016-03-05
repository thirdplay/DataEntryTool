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
' �萔
'====================================================================================================
' �J�������擾�N�G��
Private Const cstColumnDefinitionGetQuery As String = "SELECT (utc.column_id - 1) AS column_id, utc.column_name, ucc.comments, utc.data_type, CASE utc.data_type WHEN 'CHAR' THEN TO_CHAR(utc.data_length) WHEN 'VARCHAR2' THEN TO_CHAR(utc.data_length) WHEN 'NUMBER' THEN utc.data_precision || ',' || utc.data_scale ELSE ' ' END AS data_length, DECODE(utc.nullable,'N','0','1') AS is_nullable, DECODE(pkc.table_name, NULL, '0', '1') AS is_primary_key FROM user_tab_columns utc LEFT JOIN user_col_comments ucc ON(ucc.table_name = utc.table_name AND ucc.column_name = utc.column_name) LEFT JOIN (SELECT cc.table_name, cc.column_name FROM user_cons_columns cc INNER JOIN user_constraints cons ON(cons.constraint_name = cc.constraint_name AND cons.constraint_type = 'P')) pkc ON (pkc.table_name = utc.table_name AND pkc.column_name = utc.column_name) WHERE utc.table_name='${tableName}' ORDER BY utc.column_id"


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Dim mConnect As Object      ' �ڑ��I�u�W�F�N�g


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Private Sub Class_Terminate()
    Call IDataEntryDao_Disconnect
End Sub


'====================================================================================================
' DB�ڑ�
'----------------------------------------------------------------------------------------------------
' IN : setting �ݒ�
'====================================================================================================
Public Sub IDataEntryDao_Connect(setting As Setting)
On Error GoTo ErrHandler
    If mConnect Is Nothing Then
        Set mConnect = CreateObject("ADODB.Connection")

        mConnect.Open "Driver={Microsoft ODBC for Oracle};" _
            & "Server=" & setting.ServerName & ";" _
            & "UID=" & setting.UserId & ";" _
            & "PWD=" & setting.Password & ";"
        Exit Sub
    End If
ErrHandler:
    Set mConnect = Nothing
    Err.Raise Err.Number, Err.Source, "DB�ڑ����s" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' DB�ؒf
'====================================================================================================
Public Sub IDataEntryDao_Disconnect()
    If Not mConnect Is Nothing Then
        mConnect.Close
        Set mConnect = Nothing
    End If
End Sub


'====================================================================================================
' �g�����U�N�V�����J�n
'====================================================================================================
Public Sub IDataEntryDao_BeginTrans()
    Debug.Print "Oracle:BeginTrans"
End Sub


'====================================================================================================
' �R�~�b�g
'====================================================================================================
Public Sub IDataEntryDao_CommitTrans()
    Debug.Print "Oracle:CommitTrans"
End Sub


'====================================================================================================
' ���[���o�b�N
'====================================================================================================
Public Sub IDataEntryDao_RollbackTrans()
    Debug.Print "Oracle:RollbackTrans"
End Sub


'====================================================================================================
' �N�G�����s
'----------------------------------------------------------------------------------------------------
' IN : query �N�G��������
'====================================================================================================
Public Sub IDataEntryDao_Execute(query)
    Debug.Print "Execute"
End Sub


'====================================================================================================
' �w�肳�ꂽ�e�[�u���̃J������񃊃X�g���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableName �e�[�u����
' OUT: �J������񃊃X�g
'====================================================================================================
Public Function IDataEntryDao_GetColumnDefinitions(tableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim cd As ColumnDefinition
    Dim list As Collection
    Set list = New Collection

    query = Replace(cstColumnDefinitionGetQuery, "${tableName}", tableName)
    Set recordSet = mConnect.Execute(query)
    Do Until recordSet.EOF
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = recordSet("column_id")
            .ColumnName = recordSet("column_name")
            .Comments = recordSet("comments")
            .DataType = recordSet("data_type")
            .DataLength = recordSet("data_length")
            .IsNullable = recordSet("is_nullable")
            .IsPrimaryKey = recordSet("is_primary_key")
        End With
        Call list.Add(cd)
        recordSet.MoveNext
    Loop
    Set IDataEntryDao_GetColumnDefinitions = list
End Function

