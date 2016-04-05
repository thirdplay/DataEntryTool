VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseCoreOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' Oracle��Core����
'
'====================================================================================================
Implements IDatabaseCore


'====================================================================================================
' �萔
'====================================================================================================
' �J������`�擾�N�G��
Private Const cstColumnDefinitionQuery As String = "SELECT utc.column_id, utc.column_name, ucc.comments, utc.data_type, CASE utc.data_type WHEN 'CHAR' THEN TO_CHAR(utc.data_length) WHEN 'VARCHAR2' THEN TO_CHAR(utc.data_length) WHEN 'NUMBER' THEN utc.data_precision || ',' || utc.data_scale ELSE ' ' END AS data_length, DECODE(utc.nullable,'N','1','0') AS is_required, DECODE(pkc.table_name, NULL, '0', '1') AS is_primary_key FROM user_tab_columns utc LEFT JOIN user_col_comments ucc ON(ucc.table_name = utc.table_name AND ucc.column_name = utc.column_name) LEFT JOIN (SELECT cc.table_name, cc.column_name FROM user_cons_columns cc INNER JOIN user_constraints cons ON(cons.constraint_name = cc.constraint_name AND cons.constraint_type = 'P')) pkc ON (pkc.table_name = utc.table_name AND pkc.column_name = utc.column_name) WHERE utc.table_name='${tableName}' ORDER BY utc.column_id"


'====================================================================================================
' �f�[�^�x�[�X�̐ڑ���������擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xSetting �ݒ�
' OUT: �ڑ�������
'====================================================================================================
Public Function IDatabaseCore_GetConnectStr(xSetting As Setting) As String
    IDatabaseCore_GetConnectStr = "Driver={Microsoft ODBC for Oracle};" _
        & "Server=" & xSetting.ServerName & ";" _
        & "UID=" & xSetting.UserId & ";" _
        & "PWD=" & xSetting.Password & ";"
End Function


'====================================================================================================
' �J������`�擾�N�G�����擾���܂�
'----------------------------------------------------------------------------------------------------
' OUT: �J������`�擾�N�G��
'====================================================================================================
Public Function IDatabaseCore_GetColumnDefinitionQuery() As String
    IDatabaseCore_GetColumnDefinitionQuery = cstColumnDefinitionQuery
End Function


'====================================================================================================
' �f�[�^��ʂ������񂩂ǂ������肷��
'----------------------------------------------------------------------------------------------------
' IN : dataType �f�[�^���
'====================================================================================================
Public Function IDatabaseCore_IsDataTypeString(dataType As String) As Boolean
    If InStr(UCase(dataType), "CHAR") Then
        IDatabaseCore_IsDataTypeString = True
    Else
        IDatabaseCore_IsDataTypeString = False
    End If
End Function