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
Public Function GetTableInfoList(tableNames As Collection) As Collection
    Dim i As Long
    Dim ti AS TableInfo
    Dim tiList As Collection

    Set tiList = New Collection

    For i = 1 To tableNames.Count
        Set ti = New TableInfo
        ti.ColumnInfoList = mDao.GetColumnInfoList(tableNames(i))
        ti.TableName = tableNames(i)
        Call tiList.Add(ti)
    Next

    Set GetTableInfoList = tiList
End Function

