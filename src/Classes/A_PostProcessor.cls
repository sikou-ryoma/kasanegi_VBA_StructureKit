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
' このクラスモジュールの処理で管理クラスに登録されたブックは全て閉じます。
' 終了時に閉じないブックがある場合は"RemoveAndCloseByWorkbooks"で指示。
'------------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "PostProcessor"


Public Sub ExecuteFlow()

    Logger.Info "処理の開始", CLASS_NAME

    '========================================================================
    'ブックの保存など追加で処理がある場合はここに記述
    
    

    '========================================================================
    
    Call RemoveAndCloseByWorkbooks

    Logger.Info "処理の終了", CLASS_NAME

End Sub


Public Sub RemoveAndCloseByWorkbooks()

    sc.Clear    '---SheetCollectionのシートを全て解放
    
    '========================================================================
    '終了時に閉じないブックはあらかじめBookManagerからオブジェクトを解放しておく
    
    
    bm.Remove "wbMacro" '---これでこのブックが閉じなくなる
        
    
    '========================================================================
    
    bm.CloseAll '---BookManagerのブックオブジェクトを全て解放して閉じる
    
End Sub
