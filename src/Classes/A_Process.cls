VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A_Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------
' ---Process---
' ここにはフロー毎の処理を記述してください。
' このA例では各クラスへ分離していますが、必要に応じて書き換えてください。
'-----------------------------------------------

Option Explicit

Implements IProcess

Private init As A_Initializer
Private pre As A_PreProcessor
Private main As A_MainLogic
Private post As A_PostProcessor


Private Sub Class_Initialize()
    Set init = New A_Initializer
    Set pre = New A_PreProcessor
    Set main = New A_MainLogic
    Set post = New A_PostProcessor
End Sub


Private Sub IProcess_Initialize()
    init.ExecuteFlow
End Sub


Private Function IProcess_Preprocess() As Boolean
    
    IProcess_Preprocess = True
    
    If pre.ExecuteFlow = False Then
        IProcess_Preprocess = False
    End If
    
End Function


Private Sub IProcess_MainProcess()
    main.ExecuteFlow
End Sub


Private Sub IProcess_PostProcess()
    post.ExecuteFlow
End Sub
