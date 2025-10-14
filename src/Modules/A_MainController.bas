Attribute VB_Name = "A_MainController"

'------------------------------------------------------------------------
' ---MainController---
' フロー管理用メインコントローラー部です。
' 実際の処理ここでは記入せずに各モジュールにて行って下さい。
'------------------------------------------------------------------------
Option Explicit


Private Const MODULE_NAME As String = "[A_Main]"


Public Sub StartProcessing()
    
    Const PROC_NAME As String = "[MainController]"
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False


    '設定
    '------------------------------------------------------------------------
    Call A_Initialization.MainInit
    Logger.Info PROC_NAME & " 処理の開始"
    
    
    
    
    '前処理
    '------------------------------------------------------------------------
    If A_Preprocessing.RunProcessing = False Then Exit Sub
    Call WaitMsgShow
    
    
    
    
    '本処理
    '------------------------------------------------------------------------
    Call A_MainProcessing.RunProcessing
    
    
    
    
    '後処理
    '------------------------------------------------------------------------
    Call A_Postprocessing.RunProcessing
    
    
    Unload WaitMsg
    MsgBox "処理が完了しました。", vbInformation, MACRO_NAME
    Logger.Info PROC_NAME & " 正常終了"
    Application.ScreenUpdating = True
    
    Exit Sub
    
    
    
    
ErrHandler:
    
    'エラー処理
    '------------------------------------------------------------------------
    Logger.ErrorMsg PROC_NAME & " エラー発生 : " & Err.Description
    Logger.WarnMsg PROC_NAME & " 処理を中断しました"
    MsgBox "エラーが発生しました。" & vbCrLf & "エラーメッセージ : " & Err.Description, vbExclamation, MACRO_NAME
    MsgBox "処理を中断します。", vbExclamation, MACRO_NAME
    
End Sub

