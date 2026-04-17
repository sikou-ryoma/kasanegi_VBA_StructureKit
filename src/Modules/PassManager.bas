Attribute VB_Name = "PassManager"
'----------------------------------------------------------------------
' ---PassManager---
' PasswordFormを使用したユーザ用簡易パスワードフォーム用モジュール
'----------------------------------------------------------------------
Option Explicit

Private Const MODULE_NAME As String = "[PassManager]"


'---認証用フォーム
Public Function SetPassword(Optional ByVal val As String = "") As String
    
    Dim frm As PasswordForm
    Dim hasher As HashProvider
    Dim text As String
    
    Set frm = New PasswordForm
    Set hasher = New HashProvider
    
    text = val
    If text = "" Then text = "管理パスワードを入力してください。"
    
    frm.MsgLbl.Caption = text
    frm.Show
    
    If frm.Tag = "" Then
        SetPassword = ""
        Exit Function
    End If
    
    SetPassword = hasher.DJB2(frm.Tag)

End Function


'---パス変更登録フォーム
Public Sub PromptChangePassword()

    Dim inputPass As String
    Dim xmlPath As String
    Dim errMsg As String
    
    AppConfig.InitializeProject ThisWorkbook.path
    
    xmlPath = Paths.ProjectRoot & "\config\config.xml"
    
    Logger.DebugMsg MODULE_NAME & " パスワード変更開始"
    inputPass = SetPassword("現在の管理パスワードを入力してください。")
    
    If inputPass <> KANRI_PASS Then
        errMsg = "パスワード認証に失敗した"
        GoTo ExitHandler
    End If
    
    '---新規パスは2回入力を要求
    inputPass = _
        SetPassword("新しいパスワードを入力してください。" & vbCrLf & "※英数字のみ・32文字以内")
    If inputPass <> SetPassword("※ 新しいパスワードを再度入力してください。※") Then
        errMsg = "パスワードの入力が確認出来ない"
        GoTo ExitHandler
    End If
    
    If Not ConfigManager.WriteXmlValue(xmlPath, "/Config/App/Security/KanriPass", inputPass) Then
        errMsg = "xmlファイルまたはノードが存在しない"
        GoTo ExitHandler
    End If
    
    Logger.DebugMsg MODULE_NAME & " パスワード変更成功"
    MsgBox "パスワードを変更しました。", vbInformation, MACRO_NAME
    Exit Sub
    
    
ExitHandler:
    Logger.DebugMsg MODULE_NAME & " パスワード変更失敗 : " & errMsg
    MsgBox "パスワードの変更処理がキャンセルされました。" & vbCrLf & _
        "内容 : " & errMsg & "ため", vbExclamation, MACRO_NAME

End Sub



Public Sub PromptPassword_Test()

    Dim inputPass As String
    Dim xmlPath As String
    
    AppConfig.InitializeProject ThisWorkbook.path
    
    xmlPath = Paths.ProjectRoot & "\config\config.xml"
        
    inputPass = SetPassword()

    If inputPass <> KANRI_PASS Then
        MsgBox "キャンセルまたはパスワードが正しくありません。", vbExclamation, MACRO_NAME
        Exit Sub
    Else
        MsgBox "OK", vbInformation, MACRO_NAME
    End If
 
End Sub
