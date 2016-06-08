VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DmlOutputLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データ投入ロジック
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
' クエリの接頭辞
Private Const cstQuerySuffix = ";" & vbNewLine
' DMLファイル拡張子
Private Const cstFileExt = ".sql"


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDataEntryLogic As DataEntryLogic


'====================================================================================================
' コンストラクタ
'====================================================================================================
Public Sub Class_Initialize()
    Set mDataEntryLogic = New DataEntryLogic
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Public Sub Class_Terminate()
    Set mDataEntryLogic = Nothing
End Sub


'====================================================================================================
' DMLを出力します
'----------------------------------------------------------------------------------------------------
' IN : xOutputDir 出力ディレクトリ
'    : xEntryType 投入種別
'    : xEntryData 投入データ
'====================================================================================================
Public Sub OutputDml(xOutputDir As String, xEntryType As EntryType, xEntryData As EntryData)
On Error GoTo Finally
    Dim stream As ADODB.Stream
    Dim queries As Collection
    Dim query

    ' データ投入クエリの生成
    Set queries = mDataEntryLogic.MakeQueries(xEntryType, xEntryData)

    ' ストリームオープン
    Set stream = New ADODB.Stream
    stream.Mode = adModeReadWrite   ' 読み取り/書き込み権限
    stream.Type = adTypeText        'テキストデータ
    stream.Charset = "Shift_JIS"    '文字コード
    stream.Open

    For Each query In queries
        Call stream.WriteText(query & cstQuerySuffix, adWriteChar)
    Next
    stream.SaveToFile xOutputDir & xEntryData.TableName & cstFileExt, adSaveCreateOverWrite
Finally:
    ' ストリームクローズ
    stream.Close
    Set stream = Nothing

    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub
