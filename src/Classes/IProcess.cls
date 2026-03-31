VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------
' IProcess
' 各プロセス用のインターフェース
'-----------------------------------------------
Option Explicit

Public Sub Initialize()
End Sub

Public Function Preprocess() As Boolean
    Preprocess = True
End Function

Public Sub MainProcess()
End Sub

Public Sub PostProcess()
End Sub
