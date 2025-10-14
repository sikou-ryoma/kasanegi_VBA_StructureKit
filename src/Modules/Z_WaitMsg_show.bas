Attribute VB_Name = "Z_WaitMsg_show"

'----------------------------------------------------------------------------------------------------------
' ---WaitMsg_show---
' フォーム"WaitMsg"を表示
' Unloadは直接行う
'----------------------------------------------------------------------------------------------------------
Option Explicit


Public Sub WaitMsgShow(Optional ByVal seconds As Long = 2)
    
    Dim waitUntil As Date
    waitUntil = Now + TimeSerial(0, 0, seconds)
    
    With WaitMsg
        .StartUpPosition = 0
        .Left = 150
        .Top = 100
        .Show vbModeless
    End With
    
    Application.Wait waitUntil

End Sub



'---waitMsgShowの使用例

Public Sub test()

    Call WaitMsgShow(5) '引数には表示してから止める秒数を数値で指定(引数無しでは2秒)
    
    '何らかの処理
    
    Unload WaitMsg      '直接Unloadする

End Sub
