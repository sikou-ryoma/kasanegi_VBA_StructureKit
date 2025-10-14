Attribute VB_Name = "A_Postprocessing"

'------------------------------------------------------------------------
' ---Postprocessing---
' 後処理用モジュールです。
' オブジェクトの解放やブックの終了処理などを行います。
' 必要な追加処理がある場合は下記の枠内に記入してください。
'
' また、マクロの処理が完了した際に管理クラスに登録されたブックが全て閉じます。
' この時に閉じたくないブックがある場合は下記 "RemoveAndCloseByWorkbooks" の枠内より事前にクラスからオブジェクトの解放を行って下さい。
'------------------------------------------------------------------------
Option Explicit


Private Const MODULE_NAME As String = "[A_Postprocessing]"


Public Sub RunProcessing()

    Logger.Info MODULE_NAME & " 処理の開始"

    '========================================================================
    '追加で処理がある場合はここに記述
    
    

    '========================================================================
    
    Call RemoveAndCloseByWorkbooks

    Logger.Info MODULE_NAME & " 処理の終了"

End Sub


Public Sub RemoveAndCloseByWorkbooks()

    Const PROC_NAME As String = "[RemoveAndCloseByWorkbooks]"

    sc.Clear    '---SheetCollectionのシートを全て解放
    
    '========================================================================
    '終了時また中断時に閉じないブックはあらかじめBookManagerからオブジェクトを解放しておく
    
    
    
    
    If Not bm.GetWb("wbMacro") Is Nothing Then bm.Remove "wbMacro"
    
    
    
    
    '========================================================================
    
    bm.CloseAll '---BookManagerのブックオブジェクトを全て解放して閉じる
    
End Sub

