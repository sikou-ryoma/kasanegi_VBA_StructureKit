Attribute VB_Name = "LoggerManager"
'----------------------------------------------------------------------
' ---LoggerManager---
' clsLoggerのイニシャライズ用モジュールです。
' ロガーの設定をiniファイルから読み込む為のAPI宣言もここで行います。
'----------------------------------------------------------------------
Option Explicit


Public Logger As clsLogger


' INI読み取り用API宣言
'------------------------------------------------------------------------
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    
Public Function ReadIniValue(ByVal section As String, ByVal key As String, ByVal defaultVal As String, ByVal iniPath As String) As String
    Dim buffer As String * 255
    Dim ret As Long
    ret = GetPrivateProfileString(section, key, defaultVal, buffer, Len(buffer), iniPath)
    ReadIniValue = Trim(Left(buffer, ret))
End Function


Public Sub Initialize(ByVal folderPath As String)

    Const PROC_NAME As String = "[InitializeLogger]"

    Dim iniPath As String
    
    iniPath = folderPath & "\config\log_config.ini"
    
    Set Logger = New clsLogger
    With Logger
        .LogLevel = ReadIniValue("Logger", "LogLevel", "INFO", iniPath)
        .LogFolder = folderPath & "\" & ReadIniValue("Logger", "LogFolder", "log", iniPath)
        .FilePrefix = ReadIniValue("Logger", "FilePrefix", "log", iniPath)
        
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
