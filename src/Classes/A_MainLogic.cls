VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A_MainLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' ---MainProcessor---
'------------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "MainProcessor"


Public Sub ExecuteFlow()

    Logger.Info "処理の開始", CLASS_NAME
    
    '========================================================================
    'ここに処理を記述
    
    MsgBox "取得した値 : " & GetValue(sc.GetWs("selectSheet_1")), vbInformation, MACRO_NAME  '---サンプル用プロシージャ
    
    '========================================================================
        
    Logger.Info "処理の終了", CLASS_NAME

End Sub



'---サンプル用プロシージャ
Private Function GetValue(ByRef ws As SheetManager) As Variant
    
    Const PROC_NAME As String = "[GetValue]"
    
    Dim rc As Long
    
    rc = MsgBox("セルからデータを取得しますか？", vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        With ws
            Set .RNG = .sheet.Cells(1, A__)
            GetValue = .RNG.Value
            Logger.DebugMsg "取得した値 : " & .RNG.Value, PROC_NAME
        End With
    Else
        '---何もしない
    End If

End Function


