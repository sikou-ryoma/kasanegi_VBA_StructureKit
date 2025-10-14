VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------
' ---BookManager---
' 複数のブックの管理を行うクラスモジュールです。
' Dictionaryを用いてキーで管理します。
'------------------------------------------------------------------------
Option Explicit


Private Const CLASS_NAME As String = "[BookManager]"

Private wbDict As Object


'---ライブラリの呼び出しはインスタンス生成時に自動で行う
Private Sub Class_Initialize()
    Set wbDict = CreateObject("Scripting.Dictionary")
End Sub


'---ブックのディクショナリへの追加
'---ディクショナリを使用するためaddにはkeyとフルパスが必要
'---IsWorkbookOpenを呼び出して指定したブックが開いているかを判断し適切に処理
Public Sub AddWorkbook(ByVal key As String, ByVal filePath As String)

    Dim wbName As String
    wbName = Dir(filePath)
    
    Dim k As Variant
    
    For Each k In wbDict.keys
        If wbDict(k).Name = wbName Then
            MsgBox "ディクショナリ内に同じ名前のブックが存在します。", vbExclamation, MACRO_NAME
            Exit Sub
        End If
    Next k
    
    If Not wbDict.Exists(key) Then
        Dim wb As Workbook
        If IsWorkbookOpen(wbName) Then
            Set wb = Workbooks(wbName)
            wbDict.Add key, wb
        Else
            Set wb = Workbooks.Open(filePath)
            wbDict.Add key, wb
        End If
        Logger.DebugMsg CLASS_NAME & " 登録 : " & filePath & " | key : " & key
    Else
        MsgBox "キー '" & key & "' は既に存在します。", vbExclamation, MACRO_NAME
    End If
    
End Sub


'---ディクショナリよりブックの呼び出し(直接オブジェクトとして使用可能)
Public Function GetWb(ByVal key As String) As Workbook

    If wbDict.Exists(key) Then
        Set GetWb = wbDict(key)
    Else
        Set GetWb = Nothing
    End If
    
End Function


'---ディクショナリの確認用
Public Sub PrintWorkbookNames()

    Dim k As Variant
    Logger.DebugMsg CLASS_NAME & " ================================================================================="
    Logger.DebugMsg CLASS_NAME & "  BookManager内に登録されたブック"
    Logger.DebugMsg CLASS_NAME & " ---------------------------------------------------------------------------------"
    For Each k In wbDict.keys
        Debug.Print k & ": " & wbDict(k).Name
        Logger.DebugMsg CLASS_NAME & " " & k & " : " & wbDict(k).Name
    Next k
    Logger.DebugMsg CLASS_NAME & " ================================================================================="
    
End Sub


'---指定したブックをディクショナリから削除して閉じる
Public Sub CloseAndRemove(ByVal key As String, Optional ByVal saveChanges As Boolean = False)

    If wbDict.Exists(key) Then
        wbDict(key).Close saveChanges:=saveChanges
        wbDict.Remove key
        Logger.DebugMsg CLASS_NAME & " 登録ブックの終了 : " & key
    End If
    
End Sub


'---指定したブックをディクショナリから削除
Public Sub Remove(ByVal key As String)

    If wbDict.Exists(key) Then
        wbDict.Remove key
        Logger.DebugMsg CLASS_NAME & " 登録キーの削除 : " & key
    End If
    
End Sub


'---全てのブックをディクショナリから削除して閉じる
Public Sub CloseAll(Optional ByVal saveChanges As Boolean = False)

    Dim k As Variant
    
    For Each k In wbDict.keys
        wbDict(k).Close saveChanges:=saveChanges
    Next k
    
    wbDict.RemoveAll
    Logger.DebugMsg CLASS_NAME & " 登録された全ブックの終了"
    
End Sub


'---指定したブックが開いているかを確認する関数
'---AddWorkbook用の関数
Private Function IsWorkbookOpen(ByVal wbName As String) As Boolean

    Dim wb As Workbook
    
    On Error Resume Next
    Set wb = Workbooks(wbName)
    IsWorkbookOpen = Not wb Is Nothing
    On Error GoTo 0
    
End Function
