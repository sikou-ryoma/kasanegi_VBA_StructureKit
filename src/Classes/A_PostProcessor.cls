VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A_PostProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' ---PostProcessor---
' 後処理用クラスです。
' オブジェクトの解放やブックの終了処理などを行います。
' 必要な追加処理がある場合は下記の枠内に記入してください。
'
' また、このクラスモジュールの処理で管理クラスに登録されたブックは全て閉じます。
' この時に閉じたくないブックがある場合は下記 "RemoveAndCloseByWorkbooks" の枠内より事前に管理クラスからオブジェクトの解放を行うことで回避できます。
'------------------------------------------------------------------------
Option Explicit


Private Const CLASS_NAME As String = "[PostProcessor]"


Public Sub ExecuteFlow()

    Logger.Info CLASS_NAME & " 処理の開始"

    '========================================================================
    'ブックの保存など追加で処理がある場合はここに記述
    
    

    '========================================================================
    
    Call RemoveAndCloseByWorkbooks

    Logger.Info CLASS_NAME & " 処理の終了"

End Sub


Public Sub RemoveAndCloseByWorkbooks()

    Const PROC_NAME As String = "[RemoveAndCloseByWorkbooks]"

    sc.Clear    '---SheetCollectionのシートを全て解放
    
    '========================================================================
    '終了時に閉じないブックはあらかじめBookManagerからオブジェクトを解放しておく
    
    
    bm.Remove "wbMacro"
    
    
    '========================================================================
    
    bm.CloseAll '---BookManagerのブックオブジェクトを全て解放して閉じる
    
End Sub
