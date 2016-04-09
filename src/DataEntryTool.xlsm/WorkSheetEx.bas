Attribute VB_Name = "WorkSheetEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' ���[�N�V�[�g�̊g�����W���[��
'
'====================================================================================================

'====================================================================================================
' �s�f�[�^�����擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : sheetName �V�[�g��
'    : rowIndex �s�C���f�b�N�X
'    : colIndex ��C���f�b�N�X
' OUT: �s�f�[�^��
'====================================================================================================
Public Function GetRowDataCount(sheetName As String, rowIndex As Long, colIndex As Long) As Long
    Dim count As Long

    count = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex + count, colIndex).Value <> ""
            count = count + 1
        Loop
    End With
    GetRowDataCount = count
End Function


'====================================================================================================
' ��f�[�^�����擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : sheetName �V�[�g��
'    : rowIndex �s�C���f�b�N�X
'    : colIndex ��C���f�b�N�X
' OUT: ��f�[�^��
'====================================================================================================
Public Function GetColDataCount(sheetName As String, rowIndex As Long, colIndex As Long) As Long
    Dim count As Long

    count = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex, colIndex + count).Value <> ""
            count = count + 1
        Loop
    End With
    GetColDataCount = count
End Function

