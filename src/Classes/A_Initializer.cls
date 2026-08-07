VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A_Initializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' ---Initializer---
'------------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "Initializer"


Public Sub ExecuteFlow()
    
    '---汎用クラスを生成
    Set bm = New BookManager
    Set sc = New SheetCollection
    Set du = New DateUtility
    Set ctx = New ProcessStateManager
    Set context = New ProcessContext
    
    context.IsCanceled = False


    AddUsePath  '---任意のパスを登録

    'OpenFileDialog  '---ファイルを取得
    OpenMultipleFilesDialog '---複数ファイルを取得
    

    Logger.Info "設定処理の終了", CLASS_NAME

End Sub


Private Sub AddUsePath()
    
    '---マクロがThisWorkbook依存であればマネージャーへ追加
    bm.AddWorkbook "wbMacro", ThisWorkbook.FullName
    sc.AddWs bm.GetWb("wbMacro").Sheets("macro"), "wsMacro"

    '---パス管理クラスにパスを追加する
    '( Paths.SetPath "[key]", "[登録したいフォルダパス]" )
    Paths.SetPath "documents", Environ("USERPROFILE") & "\Documents"
    
    '---読み込みや書き込みを行う作業フォルダの設定
    '( Paths.SetCurrentPath "[指定したいフォルダパス]" )
    Paths.SetCurrentPath Paths.TestPath

End Sub


Private Sub OpenFileDialog()

    fs.fileName = fs.OpenFile
    If fs.fileName = False Then
        context.IsCanceled = True
        Exit Sub
    End If
    
    '---取得したファイルパスよりブックを登録
    bm.AddWorkbook "selectBook", fs.fileName
    sc.AddWs bm.GetWb("selectBook").Sheets(1), "selectSheet"

End Sub


Private Sub OpenMultipleFilesDialog()

    fs.MultipleFiles = fs.OpenMultipleFiles
    If VarType(fs.MultipleFiles) = vbBoolean Then
        context.IsCanceled = True
        Exit Sub
    End If
    
    '---取得したファイルパスより全ブックを一括登録
    '---FO.MultipleFilesの配列インデックスは1ベース
    Dim i As Long
    For i = LBound(fs.MultipleFiles) To UBound(fs.MultipleFiles)
        bm.AddWorkbook "selectBook_" & i, fs.MultipleFiles(i)
        sc.AddWs bm.GetWb("selectBook_" & i).Sheets(1), "selectSheet_1"
    Next i

End Sub
