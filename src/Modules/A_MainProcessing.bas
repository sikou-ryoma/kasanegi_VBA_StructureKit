Attribute VB_Name = "A_MainProcessing"

'------------------------------------------------------------------------
' ---MainProcessing---
' 本処理用モジュールです。
' 実際に行いたい処理は下記枠内に記述してください。
'------------------------------------------------------------------------
Option Explicit


Private Const MODULE_NAME As String = "[A_MainProcessing]"


Public Sub RunProcessing()

    Logger.Info MODULE_NAME & " 処理の開始"
    
    '========================================================================
    'ここに処理を記述またはモジュール、プロシージャをCallする
    
    
    
    MsgBox "取得した値 : " & GetValue(sc.GetWs("targetSheet")), vbInformation, MACRO_NAME  '---サンプル用プロシージャ
    
    
    
    '========================================================================
        
    Logger.Info MODULE_NAME & " 処理の終了"

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
            Logger.DebugMsg PROC_NAME & " 取得した値 : " & .RNG.Value
        End With
    Else
        '---何もしない
    End If

End Function

