VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A_PreProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' ---Preprocessor---
' 前処理用クラスです。
' 処理を行う際に必要なブックやシートの管理クラスへの事前登録を行います。
'------------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "Preprocessor"


Public Sub ExecuteFlow()
    
    If context.IsCanceled = True Then Exit Sub

    Logger.Info "処理の開始", CLASS_NAME
    
    '========================================================================
    
    '処理をここに記述
    
    '========================================================================
    
    Logger.Info "処理の終了", CLASS_NAME

End Sub


