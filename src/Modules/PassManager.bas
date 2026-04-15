Attribute VB_Name = "PassManager"
Option Explicit



Public Function SetPassword_4(Optional ByVal val As String = "") As String
    
    Dim frm As iptPass
    Dim hasher As HashProvider
    Dim text As String
    
    Set frm = New iptPass
    Set hasher = New HashProvider
    
    text = val
    If text = "" Then text = "4桁の管理パスワードを入力してください。"
    
    frm.MsgLbl.Caption = text
    frm.Show
    
    SetPassword_4 = hasher.DJB2(frm.Tag)

End Function


Public Sub PromptChangePassword()

    Dim initA As A_Initializer
    Dim inputPass As String
    Dim xmlPath As String
    
    AppConfig.InitializeProject ThisWorkbook.path
    
    Set initA = New A_Initializer
    xmlPath = Paths.ProjectRoot & "\config\config.xml"
        
    inputPass = SetPassword_4("現在の4桁の管理パスワードを入力してください。")

    If inputPass <> KANRI_PASS Then
        MsgBox "キャンセルまたはパスワードが正しくありません。", vbExclamation, MACRO_NAME
        Exit Sub
    End If
        
    inputPass = SetPassword_4("新しいパスワードを4桁で入力してください。")
    
    If Not ConfigManager.WriteXmlValue(xmlPath, "/Config/App/Security/KanriPass", inputPass) Then
        MsgBox "設定ファイルまたノードが存在しないため登録が出来ません。", vbExclamation, MACRO_NAME
    End If

End Sub

