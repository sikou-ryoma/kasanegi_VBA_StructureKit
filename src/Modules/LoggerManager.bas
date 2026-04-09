Attribute VB_Name = "LoggerManager"
'----------------------------------------------------------------------
' ---LoggerManager---
' clsLoggerのイニシャライズ用モジュールです。
'----------------------------------------------------------------------
Option Explicit

Public Logger As clsLogger


' 設定ファイルよりLogger初期化
'------------------------------------------------------------------------
Public Sub Initialize(ByVal folderPath As String)

    Const PROC_NAME As String = "[InitializeLogger]"

    Dim xmlPath As String
    xmlPath = folderPath & "\config\config.xml"
    
    Dim config As Object
    Set config = GetLoggerConfig(xmlPath)
    
    Set Logger = New clsLogger
    
    With Logger
        .LogLevel = config("LogLevel")
        .LogFolder = folderPath & "\" & config("LogFolder")
        .FilePrefix = config("FilePrefix")
        
        If Dir(.LogFolder, vbDirectory) = "" Then
            On Error Resume Next
            MkDir .LogFolder
            If Err.Number <> 0 Then
                MsgBox "ログフォルダの作成に失敗しました：" & vbCrLf & .LogFolder & vbCrLf & Err.Description, vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End If
        .Info PROC_NAME & " ================================================================================="
        .Info PROC_NAME & " ログ開始"
        .Info PROC_NAME & " ---------------------------------------------------------------------------------"
        .Info PROC_NAME & " LogLevel      : " & .LogLevel
        .Info PROC_NAME & " LogFolder     : " & .LogFolder
        .Info PROC_NAME & " FilePrefix    : " & .FilePrefix
        .Info PROC_NAME & " macro name    : " & MACRO_NAME
        .Info PROC_NAME & " macro version : " & VERSION
        .Info PROC_NAME & " ================================================================================="
    End With
    
End Sub

