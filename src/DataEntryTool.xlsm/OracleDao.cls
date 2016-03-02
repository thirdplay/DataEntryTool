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
' メンバ変数
'====================================================================================================
Dim mConnect As Object      ' 接続オブジェクト


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Call IDataEntryDao_Disconnect
End Sub


'====================================================================================================
' DB接続
'----------------------------------------------------------------------------------------------------
' IN : setting 設定
'====================================================================================================
Public Sub IDataEntryDao_Connect(setting As Setting)
On Error GoTo ErrHandler
    Set mConnect = CreateObject("ADODB.Connection")

    mConnect.Open "Driver={Microsoft ODBC for Oracle};" _
        & "Server=" & setting.ServerName & ";" _
        & "UID=" & setting.UserId & ";" _
        & "PWD=" & setting.Password & ";"
    Exit Sub
ErrHandler:
    Set mConnect = Nothing
    Err.Raise Err.Number, Err.Source, "DB接続失敗" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' DB切断
'====================================================================================================
Public Sub IDataEntryDao_Disconnect()
    If Not mConnect Is Nothing Then
        mConnect.Close
        Set mConnect = Nothing
    End If
End Sub


'====================================================================================================
' トランザクション開始
'====================================================================================================
Public Sub IDataEntryDao_BeginTrans()
    MsgBox "Oracle:BeginTrans"
End Sub


'====================================================================================================
' コミット
'====================================================================================================
Public Sub IDataEntryDao_CommitTrans()
    MsgBox "Oracle:CommitTrans"
End Sub


'====================================================================================================
' ロールバック
'====================================================================================================
Public Sub IDataEntryDao_RollbackTrans()
    MsgBox "Oracle:RollbackTrans"
End Sub


'====================================================================================================
' 指定されたテーブルのカラム情報リストを取得します。
'----------------------------------------------------------------------------------------------------
' IN : tableName テーブル名
' OUT: カラム情報リスト
'====================================================================================================
Public Function IDataEntryDao_GetColumnInfoList(tableName As String) As Collection
    Set IDataEntryDao_GetColumnInfoList = Nothing
End Function
'select
'UTC.COLUMN_NAME
',UCC.COMMENTS as COMMENT
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
' クエリ実行
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
'====================================================================================================
'Public Sub Execute(query)
'End Sub
