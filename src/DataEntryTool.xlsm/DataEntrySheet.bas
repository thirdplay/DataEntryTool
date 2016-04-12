Attribute VB_Name = "DataEntrySheet"
Option Explicit
Option Private Module

'====================================================================================================
'
' �e�[�u�������V�[�g�̃��W���[��
'
'====================================================================================================

'====================================================================================================
' �e�[�u���ݒ胊�X�g��ԋp���܂�
'----------------------------------------------------------------------------------------------------
' IN : isEntryTarget �����Ώۃt���O(True:�����Ώۂ̂�,False:�S��)
' OUT: �e�[�u���ݒ胊�X�g
'====================================================================================================
Public Function GetTableSettings(isEntryTarget As Boolean) As Object
    Dim rowIndex As Long
    Dim ts As TableSetting
    Dim isTarget As Boolean
    Dim xTableName As String
    Dim dataCount As Long
    Dim tableRange As Range
    Dim tr As Range
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    With ThisWorkbook.Worksheets(cstSheetMain)
        rowIndex = .Range(cstTableBase).Row + 1
        dataCount = WorkSheetEx.GetRowDataCount(cstSheetMain, rowIndex, TableSettingCol.PhysicsName)
        If dataCount > 0 Then
            Set tableRange = .Range(.Cells(rowIndex, TableSettingCol.PhysicsName), .Cells(rowIndex + dataCount - 1, TableSettingCol.Max))

            For Each tr In tableRange.Rows
                ' �����Ώۃt���O��true�̏ꍇ�A�����ΏۊO�̃e�[�u���͏������Ȃ�
                xTableName = tr.Cells(1, TableSettingCol.PhysicsName).Value
                isTarget = True
                If isEntryTarget Then
                    If tr.Cells(1, TableSettingCol.DataEntryTarget).Value = "" Then
                        isTarget = False
                    ElseIf Not WorkBookEx.ExistsSheet(xTableName) Then
                        Err.Raise 1000, , "�e�[�u��[" & xTableName & "]�̃V�[�g�����݂��܂���B" & vbNewLine & "�e�[�u���V�[�g�쐬���s���A�e�[�u���V�[�g���쐬���Ă��������B"
                    ElseIf ThisWorkbook.Worksheets(xTableName).Cells(cstTableRecordBase, 1).Value = "" Then
                        isTarget = False
                    End If
                End If

                If isTarget Then
                    If dic.Exists(tr.Cells(1, TableSettingCol.PhysicsName).Value) Then
                        Err.Raise 1000, , "�e�[�u��[" & xTableName & "]���d�����Ă��܂��B"
                    End If
                    Set ts = New TableSetting
                    ts.Row = tr.Row
                    ts.PhysicsName = tr.Cells(1, TableSettingCol.PhysicsName).Value
                    ts.LogicalName = tr.Cells(1, TableSettingCol.LogicalName).Value
                    ts.DataEntryTarget = tr.Cells(1, TableSettingCol.DataEntryTarget).Value
                    Call dic.Add(ts.PhysicsName, ts)
                End If
            Next
        End If
    End With
    Set GetTableSettings = dic
End Function


'====================================================================================================
' �e�[�u���ݒ�Ƀn�C�p�[�����N��ݒ肵�܂�
'----------------------------------------------------------------------------------------------------
' IN : tablSeetings �e�[�u���ݒ胊�X�g
'====================================================================================================
Public Sub SetHyperlink(tableSettings As Object)
    Dim xKey As Variant
    Dim ts As TableSetting

    With ThisWorkbook.Worksheets(cstSheetMain)
        For Each xKey In tableSettings
            Set ts = tableSettings(xKey)
            If .Cells(ts.Row, TableSettingCol.LogicalName) <> "" Then
                Call .Hyperlinks.Add(Anchor:=.Cells(ts.Row, TableSettingCol.LogicalName), Address:="#" & ts.PhysicsName & "!A1")
            End If
        Next
    End With
End Sub
