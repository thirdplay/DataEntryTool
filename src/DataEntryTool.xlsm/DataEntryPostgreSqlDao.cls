VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryPostgreSqlDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' PostgreSqlDao
'
'====================================================================================================
Implements IDataEntryDao


'====================================================================================================
' �萔
'====================================================================================================
' �J�������擾�N�G��
Private Const cstColumnInfoGetQuery As String = _
    "SELECT isc.ordinal_position - 1 AS column_id, isc.column_name, COALESCE(pd.description, '') AS comments, isc.data_type AS data_type, CASE isc.data_type WHEN 'character' THEN TO_CHAR(isc.character_maximum_length, 'FM99999999') WHEN 'character varying' THEN TO_CHAR(isc.character_maximum_length, 'FM99999999') WHEN 'numeric' THEN TO_CHAR(isc.numeric_precision, 'FM99999999') || ',' || TO_CHAR(isc.numeric_scale, 'FM99999999') ELSE '' END AS data_length, CASE isc.is_nullable WHEN 'NO' THEN '1' ELSE '0' END AS is_nullable, CASE WHEN pkc.column_name IS NOT NULL THEN '1' ELSE '0' END AS is_primary_key FROM information_schema.columns isc INNER JOIN pg_stat_all_tables psat ON( psat.schemaname = isc.table_schema AND psat.relname = isc.table_name ) INNER JOIN pg_attribute pa ON( pa.attrelid = psat.relid AND pa.attname = isc.column_name ) LEFT JOIN pg_description pd ON( pd.objoid = psat.relid AND pd.objsubid != 0 and pd.objsubid = pa.attnum ) LEFT JOIN ( SELECT tc.table_catalog, tc.table_schema, tc.table_name, " & _
    "ccu.column_name FROM information_schema.table_constraints tc INNER JOIN information_schema.constraint_column_usage ccu ON( ccu.table_catalog = tc.table_catalog AND ccu.table_schema = tc.table_schema AND ccu.table_name = tc.table_name AND ccu.constraint_name = tc.constraint_name  ) WHERE tc.constraint_type = 'PRIMARY KEY' ) pkc ON ( pkc.table_catalog = isc.table_catalog AND pkc.table_schema = isc.table_schema AND pkc.table_name = isc.table_name AND pkc.column_name = isc.column_name ) WHERE isc.table_name='${tableName}' ORDER BY isc.ordinal_position"


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
' IN : setting �f�[�^�x�[�X���
'====================================================================================================
Public Sub IDataEntryDao_Connect(setting As Setting)
On Error GoTo ErrHandler
    Set mConnect = CreateObject("ADODB.Connection")

    mConnect.Open "Driver={PostgreSQL Unicode};" _
        & "SERVER=" & setting.ServerName & ";" _
        & "DATABASE=" & setting.DatabaseName & ";" _
        & "UID=" & setting.UserId & ";" _
        & "PWD=" & setting.Password & ";" _
        & "PORT=" & setting.Port & ";"
    Exit Sub
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
    MsgBox "PostgreSql:BeginTrans"
End Sub


'====================================================================================================
' �R�~�b�g
'====================================================================================================
Public Sub IDataEntryDao_CommitTrans()
    MsgBox "PostgreSql:CommitTrans"
End Sub


'====================================================================================================
' ���[���o�b�N
'====================================================================================================
Public Sub IDataEntryDao_RollbackTrans()
    MsgBox "PostgreSql:RollbackTrans"
End Sub


'====================================================================================================
' �w�肳�ꂽ�e�[�u���̃J������񃊃X�g���擾���܂��B
'----------------------------------------------------------------------------------------------------
' IN : tableName �e�[�u����
' OUT: �J������񃊃X�g
'====================================================================================================
Public Function IDataEntryDao_GetColumnInfoList(tableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim ci As ColumnInfo
    Dim list As Collection
    Set list = New Collection

    query = Replace(cstColumnInfoGetQuery, "${tableName}", tableName)
    Set recordSet = mConnect.Execute(query)
    Do Until recordSet.EOF
        Set ci = New ColumnInfo
        With ci
            .ColumnId = recordSet("column_id")
            .ColumnName = recordSet("column_name")
            .Comments = recordSet("comments")
            .DataType = recordSet("data_type")
            .DataLength = recordSet("data_length")
            .IsNullable = recordSet("is_nullable")
            .IsPrimaryKey = recordSet("is_primary_key")
        End With
        Call list.Add(ci)
        recordSet.MoveNext
    Loop
    Set recordSet = Nothing

    Set IDataEntryDao_GetColumnInfoList = list
End Function


'====================================================================================================
' �N�G�����s
'----------------------------------------------------------------------------------------------------
' IN : query �N�G��������
'====================================================================================================
'Public Sub Execute(query)
'End Sub