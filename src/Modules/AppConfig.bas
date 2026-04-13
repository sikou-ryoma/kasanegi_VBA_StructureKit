Attribute VB_Name = "AppConfig"
'------------------------------------------------------------------------
' ---AppConfig---
' 全体の管理設定行います。
' クラスの宣言、設定ファイルより情報を読み込む関数など。
'------------------------------------------------------------------------
Option Explicit


' アプリケーション情報
'------------------------------------------------------------------------
Public VERSION As String
Public MACRO_NAME As String
Public KANRI_PASS As String
Public UTIL_GUID As String


' クラス
'------------------------------------------------------------------------
Public FO As FileOjt
Public bm As BookManager
Public sc As SheetCollection
Public du As DateUtility
Public Paths As PathConfig
Public ctx As ProcessStateManager
Public context As ProcessContext


' プロジェクト情報の初期化
'------------------------------------------------------------------------
Public Sub InitializeProject(ByVal wbPath As String)
    Dim rootPath as String

    Set FO = New FileOjt
    rootPath = FO.UpPath(wbPath)

    '---プロジェクト情報の読み込み
    AppConfig.InitializeAppConfig rootPath
    LoggerManager.Initialize rootPath

    '---プロジェクトフォルダ内の関連パスの設定
    Set Paths = New PathConfig
    Paths.init rootPath
End Sub


' 設定ファイルよりapp初期化
'------------------------------------------------------------------------
Public Sub InitializeAppConfig(ByVal folderPath As String)
    Dim xmlPath As String
    xmlPath = folderPath & "\config\config.xml"

    Dim config As Object
    Set config = GetAppConfig(xmlPath)
    
    VERSION = "v" & config("Version")
    MACRO_NAME = config("MacroName")
    KANRI_PASS = config("KanriPass")
    UTIL_GUID = config("GUID")
End Sub


