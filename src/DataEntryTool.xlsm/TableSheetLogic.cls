VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSheetLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �e�[�u���V�[�g�̃��W�b�N
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDataEntryDao As DataEntryDao


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Public Sub Class_Initialize()
    Set mDataEntryDao = DataEntryDaoFactory.GetDataEntryDao
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Public Sub Class_Terminate()
    Set mDataEntryDao = Nothing
End Sub


'====================================================================================================
' �J������`���X�g���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ�̘A�z�z��
' OUT: �J������`���X�g
'====================================================================================================
Public Function GetColumnDefinitions(tableSettings As Dictionary) As Dictionary
    Dim rs As Object
    Dim tableNames As Dictionary
    Dim tableNameInStr As String
    Dim xTableName As Variant
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim dic As Dictionary

    ' �e�[�u�������擾���A�A�z�z����쐬����
    Set rs = mDataEntryDao.GetTableName()
    Set tableNames = New Dictionary
    Do Until rs.EOF
        Call tableNames.Add(rs("table_name").Value, rs("table_name").Value)
        rs.MoveNext
    Loop

    ' �e�[�u�����̑��݃`�F�b�N
    For Each xTableName In tableSettings
        If Not tableNames.Exists(xTableName) Then
            Err.Raise ErrNumber.Error, , "�e�[�u��[" & xTableName & "]�̃J������`���擾�ł��܂���B"
        End If
    Next

    ' �J������`�̎擾
    Set rs = mDataEntryDao.GetColumnDefinition(tableSettings)

    ' �e�[�u����`�̘A���z����쐬����
    Set dic = New Dictionary
    Do Until rs.EOF
        ' �e�[�u����`����
        If Not dic.Exists(rs("table_name").Value) Then
            Set td = New TableDefinition
            td.TableName = rs("table_name").Value
            td.ColumnDefinitions = New Collection
            Call dic.Add(td.TableName, td)
        End If

        ' �J������`����
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = rs("column_id").Value
            .ColumnName = rs("column_name").Value
            .Comments = rs("comments").Value
            .DataType = rs("data_type").Value
            .DataLength = rs("data_length").Value
            .IsRequired = rs("is_required").Value
            .IsPrimaryKey = rs("is_primary_key").Value
        End With
        Call dic(rs("table_name").Value).ColumnDefinitions.Add(cd)
        rs.MoveNext
    Loop
    Set GetColumnDefinitions = dic
End Function