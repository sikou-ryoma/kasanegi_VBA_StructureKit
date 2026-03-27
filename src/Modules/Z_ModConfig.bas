Attribute VB_Name = "Z_ModConfig"

'------------------------------------------------------------------------
' ---Z_ModConfig---
' このモジュールでは全体の管理、設定行います。
' マクロ全体の定数、使用クラス、APIの宣言など。
'------------------------------------------------------------------------
Option Explicit


' マクロ名、バージョン情報の定数宣言
'------------------------------------------------------------------------
Public Const VERSION As String = "v0.9.0"
Public Const MACRO_NAME As String = "Template_Macro"
Public Const KANRI_PASS As String = "9999"  '---IptPass用のパスワード


' パブリックレベルのクラス宣言
'------------------------------------------------------------------------
Public FO As FileOjt
Public bm As BookManager
Public sc As SheetCollection
Public du As DateUtility
Public Paths As PathConfig


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


