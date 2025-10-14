Attribute VB_Name = "A_Initialization"

'------------------------------------------------------------------------
' ---Initialization---
' 設定用のモジュールです。
' ここでは各種クラスのインスタンスを生成してます。
' 基本的なパス設定はパス管理クラスの呼び出しの際に自動で設定されますが、
' 追加のパスやカレントパスの設定、そのほか追加の設定が場合は枠内に記入してください。
'------------------------------------------------------------------------
Option Explicit


Private Const MODULE_NAME As String = "[A_Initialization]"


Public Sub MainInit()

    Call ClassInit

    '========================================================================
    '追加で設定がある場合はここに記述



    '---パス管理クラスにパスを追加する
    '( Paths.SetPath "[key]", "[登録したいフォルダパス]" )
    Paths.SetPath "documents", Environ("USERPROFILE") & "\Documents"
    
    '---読み込みや書き込みを行う作業フォルダの設定
    '( Paths.SetCurrentPath "[指定したいフォルダパス]" )
    Paths.SetCurrentPath Paths.TestPath





    '========================================================================

    Logger.Info MODULE_NAME & " 設定処理の終了"

End Sub

'---※※この処理の書き換え厳禁※※

Private Sub ClassInit()

    Const PROC_NAME As String = "[ClassInit]"
        
    Set FO = New FileOjt
    
    '---clsLoggerの起動
    Call Z_LogInit.InitializeLogger(FO.UpPath(ThisWorkbook.path))
    
    Set bm = New BookManager
    Set sc = New SheetCollection
    Set du = New DateUtility
    
    '---PathConfigの設定
    Set Paths = New PathConfig
    Paths.Init FO.UpPath(ThisWorkbook.path)
    

End Sub
