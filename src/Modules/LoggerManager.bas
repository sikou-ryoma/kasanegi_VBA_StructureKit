Attribute VB_Name = "LoggerManager"
'----------------------------------------------------------------------
' ---LoggerManager---
' clsLoggerのイニシャライズ用モジュールです。
'----------------------------------------------------------------------
Option Explicit

Public Logger As LogWriter


' 設定ファイルよりLogger初期化
'------------------------------------------------------------------------
Public Sub Initialize(ByVal folderPath As String)

    Const PROC_NAME As String = "InitializeLogger"

    Dim xmlPath As String
    xmlPath = folderPath & "\config\config.xml"
    
    Dim config As Object
    Set config = GetLoggerConfig(xmlPath)
    
    Set Logger = New LogWriter
    
    With Logger
        .LogLevel = config("LogLevel")
        .LogFolder = folderPath & "\" & config("LogFolder")
        .FilePrefix = config("FilePrefix")
        
        If Dir(.LogFolder, vbDirectory) = "" Then
            On Error Resume Next
            MkDir .LogFolder
            If Err.number <> 0 Then
                MsgBox "ログフォルダの作成に失敗しました：" & vbCrLf & .LogFolder & vbCrLf & Err.Description, vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End If
        .Info "=================================================================================", PROC_NAME
        .Info "ログ開始", PROC_NAME
        .Info "---------------------------------------------------------------------------------", PROC_NAME
        .Info "LogLevel      : " & .LogLevel, PROC_NAME
        .Info "LogFolder     : " & .LogFolder, PROC_NAME
        .Info "FilePrefix    : " & .FilePrefix, PROC_NAME
        .Info "macro name    : " & MACRO_NAME, PROC_NAME
        .Info "macro version : " & VERSION, PROC_NAME
        .Info "=================================================================================", PROC_NAME
    End With
    
End Sub

