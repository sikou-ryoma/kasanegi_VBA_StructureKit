VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A_PreProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' ---Preprocessor---
' 前処理用クラスです。
' 処理を行う際に必要なブックやシートの管理クラスへの事前登録を行います。
'------------------------------------------------------------------------
Option Explicit


Private Const CLASS_NAME As String = "[Preprocessor]"


Public Sub ExecuteFlow()

    Logger.Info CLASS_NAME & " 処理の開始"
    context.IsCanceled = False

    '========================================================================
    '事前処理などを記述してください。
    
    
        
    
    '---FO.OpFileでファイル選択ダイアログを開き取得したファイルパスをFO.FileNmに保持
    '---ファイル選択ダイアログはキャンセルでFalseを返すので中断処理を挟む
    '------------------------------------------------------------------------
'    FO.FileNm = FO.OpFile
'    If FO.FileNm = False Then
'        context.IsCanceled = True
'        Exit Sub
'    End If
    
    '---ファイル選択ダイアログより取得したファイルパスよりブックを登録
'    bm.AddWorkbook "selectBook", FO.FileNm
'    sc.AddWs bm.GetWb("selectBook").Sheets(1), "selectSheet"
    
    
    
    
    '---FO.OpMultipleFilesで選択ダイアログを開き取得した複数のファイルパスをFO.MultipleFilesに配列で保持
    '---ファイル選択ダイアログはキャンセルでFalseを返すので中断処理を挟む
    '------------------------------------------------------------------------
    FO.MultipleFiles = FO.OpMultipleFiles
    If VarType(FO.MultipleFiles) = vbBoolean Then
        context.IsCanceled = True
        Exit Sub
    End If
    
    '---ファイル選択ダイアログより取得したファイルパスより全ブックを一括登録
    '---FO.MultipleFilesの配列インデックスは1ベースのため注意
    Dim i As Long
    For i = LBound(FO.MultipleFiles) To UBound(FO.MultipleFiles)
        bm.AddWorkbook "selectBook_" & i, FO.MultipleFiles(i)
        sc.AddWs bm.GetWb("selectBook_" & i).Sheets(1), "selectSheet_" & i
    Next i
    

    '========================================================================
    
    Logger.Info CLASS_NAME & " 処理の終了"

End Sub


