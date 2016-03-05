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
Private mDao As IDataEntryDao   ' データ投入Daoインターフェース


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Set mDao = Nothing
End Sub


'====================================================================================================
' モデルを構成します
'----------------------------------------------------------------------------------------------------
' IN : setting 設定
'====================================================================================================
Public Sub Setup(setting As Setting)
    Set mDao = DaoFactory.Create(setting)
    Call mDao.Connect(setting)
End Sub


'====================================================================================================
' テーブル名一覧のテーブル情報をDBから取得し、返却します
'----------------------------------------------------------------------------------------------------
' IN : tableNames テーブル名リスト
' OUT: テーブル情報リスト
'====================================================================================================
Public Function GetTableDefinitions(tableNames As Collection) As Collection
    Dim i As Long
    Dim td AS TableDefinition
    Dim list As Collection

    Set list = New Collection

    For i = 1 To tableNames.Count
        Set td = New TableDefinition
        td.ColumnDefinitions = mDao.GetColumnDefinitions(tableNames(i))
        td.TableName = tableNames(i)
        Call list.Add(td)
    Next

    Set GetTableDefinitions = list
End Function

