Attribute VB_Name = "Z_OpenFiles"
Option Explicit

Private Const MODULE_NAME As String = "[OpenFiles]"



'---フルパスで指定したブックを開く処理(開いている場合はアクティブにする)
Private Sub FileOpenOrActivate(ByVal fullPath As String)

    Dim wb As Workbook
    Dim fileName As String
    Dim flag As Boolean
    
    fileName = Dir(fullPath)
    flag = False
    
    For Each wb In Workbooks
        If wb.Name = fileName Then
            flag = True
            Exit For
        End If
    Next wb
    
    If flag Then
        wb.Activate
        Logger.Info MODULE_NAME & " ブックをアクティブ化 : " & wb.Name
    Else
        Workbooks.Open fullPath
        Logger.Info MODULE_NAME & " ブックの起動 : " & fullPath
    End If

End Sub


'---パスワード要求付き
Public Sub TargetBook()

    A_Initialization.MainInit
    
    Dim frm As New iptPass
    Dim inputPass As String
    
    frm.Show
    inputPass = frm.Tag
        
    If inputPass = KANRI_PASS Then
        Call FileOpenOrActivate(Paths.TestPath & "\SampleData\TestBook.xlsx")
    Else
        MsgBox "キャンセルまたはパスワードが正しくありません。", vbExclamation, MACRO_NAME
    End If

    Unload WaitMsg

End Sub
