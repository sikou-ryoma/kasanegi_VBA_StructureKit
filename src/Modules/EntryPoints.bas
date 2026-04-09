Attribute VB_Name = "EntryPoints"
'------------------------------------------------------------------------
' ---EntryPoints---
' アプリケーション起動用エントリーポイントです。
' 必ずここを起点にしてコントロールボタンなどに割り当ててください。
'------------------------------------------------------------------------
Option Explicit


Private Const MODULE_NAME As String = "[EntryPoints]"


Public Sub Run_A()

    Dim app As New ApplicationService
    Dim proc As New A_Process
    
    AppConfig.InitializeProject ThisWorkbook.path
    
    app.ExecuteApp proc
        
End Sub
