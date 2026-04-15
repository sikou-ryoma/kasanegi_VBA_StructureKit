Attribute VB_Name = "ConfigManager"
'----------------------------------------------------------------------
' ---ConfigManager---
' configファイルよりを各設定を読み込む関数等をまとめています。
'----------------------------------------------------------------------
Option Explicit



' INI読み取り用API宣言 (必要に応じて使用)
'------------------------------------------------------------------------
'Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
'    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
'    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'
'Public Function ReadIniValue(ByVal section As String, ByVal key As String, ByVal defaultVal As String, ByVal iniPath As String) As String
'    Dim buffer As String * 255
'    Dim ret As Long
'    ret = GetPrivateProfileString(section, key, defaultVal, buffer, Len(buffer), iniPath)
'    ReadIniValue = Trim(Left(buffer, ret))
'End Function


' XML読み取り用関数
'------------------------------------------------------------------------
Public Function ReadXmlValue(ByVal xmlPath As String, ByVal xpath As String) As String
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.Load xmlPath
    If xmlDoc.parseError.ErrorCode <> 0 Then
        ReadXmlValue = ""
        Exit Function
    End If
    Dim node As Object
    Set node = xmlDoc.selectSingleNode(xpath)
    If Not node Is Nothing Then
        ReadXmlValue = node.text
    Else
        ReadXmlValue = ""
    End If
End Function


' XML書き込み用関数
'------------------------------------------------------------------------
Public Function WriteXmlValue(ByVal xmlPath As String, ByVal xpath As String, ByVal value As String) As Boolean
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.Load xmlPath
    If xmlDoc.parseError.ErrorCode <> 0 Then
        WriteXmlValue = False
        Exit Function
    End If
    Dim node As Object
    Set node = xmlDoc.selectSingleNode(xpath)
    If Not node Is Nothing Then
        node.text = value
        xmlDoc.Save xmlPath
        WriteXmlValue = True
    Else
        WriteXmlValue = False
    End If
End Function


' XMLよりアプリケーション設定を読み込む関数
'------------------------------------------------------------------------
Public Function GetAppConfig(ByVal xmlPath As String) As Object
    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")
    config("Version") = ReadXmlValue(xmlPath, "/Config/App/Meta/Version")
    config("MacroName") = ReadXmlValue(xmlPath, "/Config/App/Meta/MacroName")
    config("KanriPass") = ReadXmlValue(xmlPath, "/Config/App/Security/KanriPass")
    config("GUID") = ReadXmlValue(xmlPath, "/Config/App/Security/GUID")
    Set GetAppConfig = config
End Function


' XMLよりロガーの設定を読み込む関数
'------------------------------------------------------------------------
Public Function GetLoggerConfig(ByVal xmlPath As String) As Object
    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")
    config("LogLevel") = ReadXmlValue(xmlPath, "/Config/Logger/LogLevel")
    config("LogFolder") = ReadXmlValue(xmlPath, "/Config/Logger/LogFolder")
    config("FilePrefix") = ReadXmlValue(xmlPath, "/Config/Logger/FilePrefix")
    Set GetLoggerConfig = config
End Function

