Attribute VB_Name = "DataEntryTool"
Option Explicit


'====================================================================================================
' テーブルシートを作成します
'====================================================================================================
Public Sub CreateTableSheet()
On Error GoTo Finally
    Dim control As DataEntryControl
    Dim start As Single
    Dim finish As Single
    start = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .CreateTableSheet
    End With

Finally:
    finish = Timer
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox "テーブルシートの作成に失敗しました" & vbNewLine & Err.Description
    Else
        Debug.Print "処理時間:" & (finish - start)
        MsgBox "テーブルシートの作成が完了しました" & ":" & (finish - start)
    End If
End Sub


'====================================================================================================
' データを登録します
'====================================================================================================
Public Sub RegisterData()
    Call ExecuteEntryData(EntryType.Register)
End Sub


'====================================================================================================
' データを更新します
'====================================================================================================
Public Sub UpdateData()
    Call ExecuteEntryData(EntryType.Update)
End Sub


'====================================================================================================
' データを削除します
'====================================================================================================
Public Sub RemoveData()
    Call ExecuteEntryData(EntryType.Remove)
End Sub


'====================================================================================================
' データ投入の実行
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'====================================================================================================
Private Sub ExecuteEntryData(xEntryType As EntryType)
On Error GoTo Finally
    Dim operationStr As String
    Dim control As DataEntryControl
    Dim start As Single
    Dim finish As Single
    start = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .EntryData(xEntryType)
    End With

Finally:
    finish = Timer
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    operationStr = GetOperationStr(xEntryType)
    If Err.Number <> 0 Then
        MsgBox "データ" & operationStr & "に失敗しました" & vbNewLine & Err.Description
    Else
        Debug.Print "処理時間:" & (finish - start)
        MsgBox "データ" & operationStr & "が完了しました" & ":" & (finish - start)
    End If
End Sub


'====================================================================================================
' 投入種類に対応する操作文字列を取得します
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
' OUT: 操作文字列
'====================================================================================================
Private Function GetOperationStr(xEntryType As EntryType)
    Select Case xEntryType
        Case EntryType.Register
            GetOperationStr = "登録"
        Case EntryType.Register
            GetOperationStr = "更新"
        Case EntryType.Register
            GetOperationStr = "削除"
    End Select
End Function

