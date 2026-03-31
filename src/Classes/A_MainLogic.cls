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
' 本処理用クラスです。
' 実際に行いたい処理は下記枠内に記述してください。
'------------------------------------------------------------------------
Option Explicit


Private Const CLASS_NAME As String = "[MainProcessor]"


Public Sub ExecuteFlow()

    Logger.Info CLASS_NAME & " 処理の開始"
    
    '========================================================================
    'ここに処理を記述またはモジュール、プロシージャをCallする
    
    MsgBox "取得した値 : " & GetValue(sc.GetWs("selectSheet_1")), vbInformation, MACRO_NAME  '---サンプル用プロシージャ
    
    '========================================================================
        
    Logger.Info CLASS_NAME & " 処理の終了"

End Sub



'---サンプル用プロシージャ
Private Function GetValue(ByRef ws As SheetManager) As Variant
    
    Const PROC_NAME As String = "[GetValue]"
    
    Dim rc As Long
    
    rc = MsgBox("セルからデータを取得しますか？", vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        With ws
            Set .RNG = .sheet.Cells(1, A__)
            GetValue = .RNG.value
            Logger.DebugMsg PROC_NAME & " 取得した値 : " & .RNG.value
        End With
    Else
        '---何もしない
    End If

End Function


