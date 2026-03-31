VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "[ApplicationService]"


Public Sub ExecuteApp(ByVal proc As IProcess)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    proc.Initialize
    
    If proc.Preprocess = False Then Exit Sub
    WaitMsgShow
    
    proc.MainProcess
    
    proc.PostProcess
    
    Unload WaitMsg
    MsgBox "処理が完了しました。", vbInformation, MACRO_NAME
    Logger.Info CLASS_NAME & " 正常終了"
    Application.ScreenUpdating = True
    
    Exit Sub
    
    
    
ErrHandler:
    
    Logger.ErrorMsg CLASS_NAME & " エラー発生 : " & Err.Description
    Logger.WarnMsg CLASS_NAME & " 処理を中断しました"
    MsgBox "エラーが発生しました。" & vbCrLf & "エラーメッセージ : " & Err.Description, vbExclamation, MACRO_NAME
    Unload WaitMsg
    MsgBox "処理を中断します。", vbExclamation, MACRO_NAME
    proc.PostProcess
    
End Sub

