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

Private Const CLASS_NAME As String = "ApplicationService"


Public Sub ExecuteApp(ByVal proc As IProcess)

    'On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    proc.Initialize
    
    proc.Preprocess
    
    If proc.CanExecute = False Then Exit Sub
    
    'WaitMsgShow
    ShowProgress
    progFrm.Value 20
        
    proc.MainProcess
    progFrm.Value 70
    
    progFrm.Value 99
    
    Application.ScreenUpdating = True
        
    progFrm.SelfClose
    'Unload WaitMsg
    
    proc.PostProcess
    
    MsgBox "処理が完了しました。", vbInformation, MACRO_NAME
    Logger.Info "正常終了", CLASS_NAME
    
    Exit Sub
    
    
ErrHandler:

    progFrm.SelfClose
    'Unload WaitMsg
    Logger.ErrorMsg "エラー発生 : " & Err.Description, CLASS_NAME
    Logger.WarnMsg "処理を中断しました", CLASS_NAME
    MsgBox "エラーが発生しました。" & vbCrLf & "エラーメッセージ : " & Err.Description, vbExclamation, MACRO_NAME
    Unload WaitMsg
    MsgBox "処理を中断します。", vbExclamation, MACRO_NAME
    proc.PostProcess
    
End Sub

