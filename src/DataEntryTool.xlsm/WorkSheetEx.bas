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
    Dim dataCnt As Long

    dataCnt = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex + dataCnt, colIndex).Value <> ""
            dataCnt = dataCnt + 1
        Loop
    End With
    GetRowDataCount = dataCnt
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
    Dim dataCnt As Long

    dataCnt = 0
    With ThisWorkbook.Worksheets(sheetName)
        Do While .Cells(rowIndex, colIndex + dataCnt).Value <> ""
            dataCnt = dataCnt + 1
        Loop
    End With
    GetColDataCount = dataCnt
End Function

