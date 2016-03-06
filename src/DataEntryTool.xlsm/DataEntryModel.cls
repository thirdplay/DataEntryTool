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
Public Sub Setup(Setting As Setting)
    Set mDao = DaoFactory.Create(Setting)
    Call mDao.Connect(Setting)
End Sub


'====================================================================================================
' テーブル設定リストのテーブル情報をDBから取得し、返却します
'----------------------------------------------------------------------------------------------------
' IN : tableSetting テーブル設定リスト
' OUT: テーブル定義リスト
'====================================================================================================
Public Function GetTableDefinitions(tableSetting As Collection) As Collection
    Dim ts As TableSetting
    Dim td As TableDefinition
    Dim list As Collection

    Set list = New Collection

    For Each ts In tableSetting
        Set td = New TableDefinition
        td.ColumnDefinitions = mDao.GetColumnDefinitions(ts.PhysicsName)
        td.TableName = ts.PhysicsName
        Call list.Add(td)
    Next

    Set GetTableDefinitions = list
End Function

