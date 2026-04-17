Attribute VB_Name = "PassManager"
'----------------------------------------------------------------------
' ---PassManager---
' IptPassを使用したユーザ用簡易パスワードフォーム用モジュール
'----------------------------------------------------------------------
Option Explicit


'---認証用フォーム
Public Function SetPassword_4(Optional ByVal val As String = "") As String
    
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
        SetPassword_4 = ""
        Exit Function
    End If
    
    SetPassword_4 = hasher.DJB2(frm.Tag)

End Function


'---認証→変更登録フォーム
Public Sub PromptChangePassword()

    Dim inputPass As String
    Dim xmlPath As String
    
    AppConfig.InitializeProject ThisWorkbook.path
    
    xmlPath = Paths.ProjectRoot & "\config\config.xml"
        
    inputPass = SetPassword_4("現在の管理パスワードを入力してください。")

    If inputPass <> KANRI_PASS Then
        MsgBox "キャンセルまたはパスワードが正しくありません。", vbExclamation, MACRO_NAME
        Exit Sub
    End If
        
    inputPass = SetPassword_4("新しく設定するパスワードを入力してください。" & vbCrLf & "※英数字のみ・32文字以内")
    
    If inputPass = "" Then
        MsgBox "入力が確認出来ません。", vbExclamation, MACRO_NAME
        Exit Sub
    End If
    
    If Not ConfigManager.WriteXmlValue(xmlPath, "/Config/App/Security/KanriPass", inputPass) Then
        MsgBox "設定ファイルまたノードが存在しないため登録が出来ません。", vbExclamation, MACRO_NAME
    End If

End Sub

