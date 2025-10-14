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
    
    
    
    Call GetValue("セルからデータを取得しますか？") '---サンプル用プロシージャ
    
    
    
    '========================================================================
        
    Logger.Info MODULE_NAME & " 処理の終了"

End Sub



'---サンプル用プロシージャ
Private Sub GetValue(ByVal message As String)
    
    Dim rc As Long
    
    rc = MsgBox(message, vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        With sc.GetWs("targetSheet")
            Set .RNG = .sheet.Cells(1, A__)
            MsgBox "取得した値 : " & .RNG.Value, vbInformation, MACRO_NAME
            Logger.DebugMsg MACRO_NAME & " 取得した値 : " & .RNG.Value
        End With
    Else
        '---何もしない
    End If

End Sub

