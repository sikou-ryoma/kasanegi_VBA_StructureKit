VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' ---PathConfig---
' プロジェクトフォルダ内の主要フォルダパスやそのほか複数のパス管理を行うクラスモジュールです。
' 主要なパスはプロパティで管理、追加の登録はDictionaryとキーで管理します。
' ※インスタンスの生成後は必ず.Initを行うこと。
'------------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "PathConfig"

Private mProjectRoot As String
Private mReportsPath As String
Private mTempPath As String
Private mTestPath As String
Private mFormatPath As String
Private mCustomPaths As Object '---管理にディクショナリを使用


Private Sub Class_Initialize()

    mProjectRoot = ThisWorkbook.path
    Set mCustomPaths = CreateObject("Scripting.Dictionary")

End Sub


'---固定系パス
Public Property Get ProjectRoot() As String
    ProjectRoot = mProjectRoot
End Property

Public Property Get ReportsPath() As String
    ReportsPath = mReportsPath
End Property

Public Property Get TempPath() As String
    TempPath = mTempPath
End Property

Public Property Get TestPath() As String
    TestPath = mTestPath
End Property

Public Property Get FormatPath() As String
    FormatPath = mFormatPath
End Property


'---固定パスのイニシャライズ
'---Initの呼び出しは必ず行う
Public Sub init(ByVal basePath As String)
    
    mProjectRoot = basePath
    mReportsPath = mProjectRoot & "\output\reports"
    mTempPath = mProjectRoot & "\output\temp"
    mTestPath = mProjectRoot & "\test"
    mFormatPath = mProjectRoot & "\workbook\format"
    
    Logger.DebugMsg "=================================================================================", CLASS_NAME
    Logger.DebugMsg "各パスの設定", CLASS_NAME
    Logger.DebugMsg "---------------------------------------------------------------------------------", CLASS_NAME
    Logger.DebugMsg "ProjectRoot : " & mProjectRoot, CLASS_NAME
    Logger.DebugMsg "ReportsPath : " & mReportsPath, CLASS_NAME
    Logger.DebugMsg "TempPath    : " & mTempPath, CLASS_NAME
    Logger.DebugMsg "TestPath    : " & mTestPath, CLASS_NAME
    Logger.DebugMsg "FormatPath  : " & mFormatPath, CLASS_NAME
    Logger.DebugMsg "=================================================================================", CLASS_NAME
    
End Sub


'---カスタムパス
Public Sub SetPath(ByVal key As String, ByVal path As String)
    If mCustomPaths.Exists(key) Then
        mCustomPaths(key) = path
        Logger.DebugMsg "パスを追加 : " & path & " | key : " & key, CLASS_NAME
    Else
        mCustomPaths.Add key, path
        Logger.DebugMsg "パスを追加 : " & path & " | key : " & key, CLASS_NAME
    End If
End Sub

Public Function GetPath(ByVal key As String) As String
    If mCustomPaths.Exists(key) Then
        GetPath = mCustomPaths(key)
    Else
        Err.Raise 9999, , "Path for key '" & key & "' is not set."
    End If
End Function


'---作業用ドライブとパスの設定
'---引数でパスを指定してください。指定が無い場合はプロジェクトルートフォルダが指定されます。
Public Sub SetCurrentPath(Optional ByVal basePath As String = "")

    If basePath = "" Then basePath = mProjectRoot
    
    ChDrive Left(basePath, 1)
    ChDir basePath
    Logger.DebugMsg "カレントパスを設定 : " & basePath, CLASS_NAME

End Sub

