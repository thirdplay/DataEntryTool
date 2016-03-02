VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データ投入モデル
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mSetting As Setting     ' データベース情報


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Set mSetting = Nothing
End Sub


'====================================================================================================
' 設定の設定
'====================================================================================================
Public Property Let Setting(setting As Setting)
    Set mSetting = setting
End Property


'====================================================================================================
' テーブル名一覧のテーブル情報をDBから取得し、返却します
'----------------------------------------------------------------------------------------------------
' IN : tableNames テーブル名リスト
' OUT: テーブル情報リスト
'====================================================================================================
Public Function GetTableInfoList(tableNames As Collection) As Collection
On Error GoTo Finally
    Dim i As Long
    Dim dao As IDataEntryDao
    Dim ti AS TableInfo
    Dim tiList As Collection

    Set tiList = New Collection
    Set dao = DaoFactory.Create(mSetting.DatabaseType)

    Call dao.Connect(mSetting)
    For i = 1 To tableNames.Count
        Set ti = New TableInfo
        ti.ColumnInfoList = dao.GetColumnInfoList(tableNames(i))
        ti.TableName = tableNames(i)
        Call tiList.Add(ti)
    Next
    Call dao.Disconnect

    Set GetTableInfoList = tiList
Finally:
    Set dao = Nothing
    Set ti = Nothing
    Set tiList = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

