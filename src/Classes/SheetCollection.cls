VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------
' ---SheetCollection---
' 複数の SheetManager インスタンスを一括管理するためのクラス
' 各ワークシートごとに既存の SheetManager クラスを使ってインスタンスを生成
' 名前(キー)でアクセス、追加、削除が可能
' 全体に対するメソッド呼び出し(CallByName)をサポート
'
' 前提: プロジェクトにクラスモジュール "SheetManager" が存在すること
'------------------------------------------------------------------------
Option Explicit


Private Const CLASS_NAME As String = "[SheetCollection]"

Private pDict As Object


'---クラス初期化
Private Sub Class_Initialize()
    
    Set pDict = CreateObject("Scripting.Dictionary")

End Sub


'---クラス解放
Private Sub Class_Terminate()
    
    On Error Resume Next
    Set pDict = Nothing
    On Error GoTo 0

End Sub


'---ワークシートから新しい SheetManager を生成して追加
'---既に同じキーがある場合はエラーを発生させる
Public Sub AddWs(ByVal ws As Worksheet, Optional ByVal key As String)

    Dim sm As SheetManager
    If ws Is Nothing Then Err.Raise vbObjectError + 1000, "SheetCollection.Add", "Worksheet が無効です"

    If key = "" Then key = ws.Name
    If pDict.Exists(key) Then Err.Raise vbObjectError + 1001, "SheetCollection.Add", "キー '" & key & "' は既に存在します"

    Set sm = New SheetManager
    Set sm.sheet = ws
    pDict.Add key, sm
    Logger.DebugMsg CLASS_NAME & " 登録 : " & sm.sheet.Name & " | key : " & key
    
End Sub


'---指定したシートをディクショナリから削除する
Public Sub Remove(ByVal key As String)
    
    If pDict.Exists(key) Then
        pDict.Remove key
        Logger.DebugMsg CLASS_NAME & " 登録キーの削除 : " & key
    End If
    
End Sub


'---キーから SheetManager を返す(存在しなければ Nothing)
Public Function GetWs(ByVal key As String) As SheetManager
    If pDict.Exists(key) Then
        Set GetWs = pDict(key)
    Else
        Set GetWs = Nothing
    End If
End Function


'---インデックスで取得(For i = 1 To Count の想定)
Public Function GetByIndex(ByVal idx As Long) As SheetManager
    If idx < 1 Or idx > pDict.Count Then
        Set GetByIndex = Nothing
    Else
        Dim k As Variant
        k = pDict.keys()(idx - 1) '---Keys は 0-based の配列を返す
        Set GetByIndex = pDict(k)
    End If
End Function


'---登録数
Public Property Get Count() As Long
    Count = pDict.Count
End Property


'---登録されているキーの配列を返す
Public Property Get keys() As Variant
    If pDict.Count = 0 Then
        keys = Array()
    Else
        keys = pDict.keys
    End If
End Property


'---登録されている SheetManager インスタンスを Collection として返す
Public Function Items() As Collection
    Dim coll As New Collection
    Dim key As Variant
    For Each key In pDict.keys
        coll.Add pDict(key)
    Next key
    Set Items = coll
End Function


'---指定ブック内の全ワークシートを読み込み(既存キーはスキップ)
Public Sub LoadAll(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        On Error Resume Next
        If Not pDict.Exists(ws.Name) Then Add ws, ws.Name
        On Error GoTo 0
    Next ws
End Sub


'---各 SheetManager に対してメソッドを一括実行
Public Sub ForEachCall(ByVal methodName As String, ParamArray args() As Variant)
    Dim key As Variant
    Dim obj As Object
    For Each key In pDict.keys
        Set obj = pDict(key)
        If Not obj Is Nothing Then
            '---CallByName でメソッド呼び出し
            On Error Resume Next
            Select Case UBound(args)
                Case -1
                    CallByName obj, methodName, VbMethod
                Case 0
                    CallByName obj, methodName, VbMethod, args(0)
                Case 1
                    CallByName obj, methodName, VbMethod, args(0), args(1)
                Case 2
                    CallByName obj, methodName, VbMethod, args(0), args(1), args(2)
                Case 3
                    CallByName obj, methodName, VbMethod, args(0), args(1), args(2), args(3)
                Case Else
                    '---4つ以上の引数が必要な場合は適宜拡張
                    CallByName obj, methodName, VbMethod, args
            End Select
            If Err.Number <> 0 Then
                '---エラーは無視して次へ（必要ならログ出力する）
                Debug.Print "ForEachCall error on key=" & key & " method=" & methodName & " Err=" & Err.Number & ":" & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next key
End Sub


'---キーの存在確認
Public Function Exists(ByVal key As String) As Boolean

    Exists = pDict.Exists(key)

End Function


'---すべてのシートの登録を削除(各インスタンスの解放)
Public Sub Clear()
    Dim key As Variant
    For Each key In pDict.keys
        On Error Resume Next
        Set pDict(key) = Nothing
        On Error GoTo 0
    Next key
    pDict.RemoveAll
    Logger.DebugMsg CLASS_NAME & " 登録された全シートの各インスタンス解放"
End Sub

