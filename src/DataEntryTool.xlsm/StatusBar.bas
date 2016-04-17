Attribute VB_Name = "StatusBar"
Option Explicit
Option Private Module

'====================================================================================================
'
' ステータスバーのモジュール
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
' ステータスバーの書式
Private Const cstStatusBarFormat = "${processName} ${progressRate}% ${progressBar}"


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDisplayStatusBar As Boolean    ' ステータスバー表示フラグ(退避用)
Private mIsDisplay As Boolean           ' 表示フラグ
Private mProcessName As String          ' プロセス名
Private mProgressCnt As Long            ' 進捗カウント
Private mProgressMax As Long            ' 進捗最大カウント
Private mStatusBarContens As String     ' ステータスバーの内容


'====================================================================================================
' ステータスバーに進捗状況を表示します
' ---------------------------------------------------------------------------------------------------
' IN : processName プロセス名
'    : progressMax 進捗カウントの最大値
'    : progressCnt 初期進捗カウント
'====================================================================================================
Public Sub ShowProgress(ByVal processName As String, Optional ByVal progressMax As Long = 100, Optional progressCnt As Long = 0)
    If Not mIsDisplay Then
        ' ステータスバーの表示
        mDisplayStatusBar = Application.DisplayStatusBar
        Application.DisplayStatusBar = True

        ' 進捗状況の初期化
        mIsDisplay = True
        mProcessName = processName
        mProgressMax = progressMax
        mProgressCnt = 0

        ' 進捗状況の増加
        Call IncreaseProgress(progressCnt)
    End If
End Sub


'====================================================================================================
' ステータスバーを非表示にします
'====================================================================================================
Public Sub Hide()
    If mIsDisplay Then
        mIsDisplay = False
        mProcessName = ""
        mProgressCnt = 0
        mProgressMax = 0
        mStatusBarContens = ""

        ' ステータスバーの表示設定を復元する
        Application.StatusBar = False
        Application.DisplayStatusBar = mDisplayStatusBar
    End If
End Sub


'====================================================================================================
' 進捗状況を増やします
' ---------------------------------------------------------------------------------------------------
' IN : progressCnt 進捗カウント
'====================================================================================================
Public Sub IncreaseProgress(Optional ByVal progressCnt As Long = 1)
    Dim statusBarContens As String
    Dim progressRate As Byte
    Dim progressStatus As Byte

    ' 表示中以外は処理しない'
    If Not mIsDisplay Then
        Exit Sub
    End If

    ' 進捗率の更新
    mProgressCnt = mProgressCnt + progressCnt
    If mProgressCnt > mProgressMax Then
        mProgressCnt = mProgressMax
    End If

    ' 進捗状況の作成
    progressRate = CByte(mProgressCnt / mProgressMax * 100)
    progressStatus = progressRate / 10
    statusBarContens = cstStatusBarFormat
    statusBarContens = Replace(statusBarContens, "${processName}", mProcessName)
    statusBarContens = Replace(statusBarContens, "${progressRate}", progressRate)
    statusBarContens = Replace(statusBarContens, "${progressBar}", String(progressStatus, "■") & String(10 - progressStatus, "□"))

    ' ステータスバーの内容に変化がある場合、更新
    If statusBarContens <> mStatusBarContens Then
        Application.StatusBar = statusBarContens
        mStatusBarContens = statusBarContens
    End If
End Sub
