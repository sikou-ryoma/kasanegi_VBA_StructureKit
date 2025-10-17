Attribute VB_Name = "A_Preprocessing"

'------------------------------------------------------------------------
' ---Preprocessing---
' 前処理用モジュールです。
' 処理を行う際に必要なブックやシートの管理クラスへの事前登録を行います。
'------------------------------------------------------------------------
Option Explicit


Private Const MODULE_NAME As String = "[A_Preprocessing]"


Public Function RunProcessing() As Boolean

    Logger.Info MODULE_NAME & " 処理の開始"
    RunProcessing = True

    '========================================================================
    'ブック及びシートの管理クラスへの登録などの事前処理を記述してください。
    
    
    
    bm.AddWorkbook "wbMacro", ThisWorkbook.FullName
    sc.AddWs bm.GetWb("wbMacro").Sheets("macro"), "wsMacro"
    
    '---FO.OpFileでファイル選択ダイアログを開き取得したファイルパスをFO.FileNmに保持
    '---ファイル選択ダイアログはキャンセルでFalseを返すので中断処理を挟む
'    FO.FileNm = FO.OpFile
'    If FO.FileNm = False Then
'        RunProcessing = False   '---ファイル選択ダイアログのキャンセル時、関数は必ずFalseを返しておく
'        Logger.Info MODULE_NAME & " キャンセルのため処理を中断"
'        Exit Function
'    End If
    
    '---ファイル選択ダイアログより取得したファイルパスよりブックを登録
'    bm.AddWorkbook "selectBook", FO.FileNm
'    sc.AddWs bm.GetWb("selectBook").Sheets(1), "selectSheet"
    
    
    
    '---FO.OpMultipleFilesで選択ダイアログを開き取得した複数のファイルパスをFO.MultipleFilesに配列で保持
    '---ファイル選択ダイアログはキャンセルでFalseを返すので中断処理を挟む
    FO.MultipleFiles = FO.OpMultipleFiles
    If VarType(FO.MultipleFiles) = vbBoolean Then
        RunProcessing = False   '---ファイル選択ダイアログのキャンセル時、関数は必ずFalseを返しておく
        Logger.Info MODULE_NAME & " キャンセルのため処理を中断"
        Exit Function
    End If
    
    '---ファイル選択ダイアログより取得したファイルパスより全ブックを一括登録
    '---FO.MultipleFilesの配列インデックスは1ベース
    Dim i As Long
    For i = LBound(FO.MultipleFiles) To UBound(FO.MultipleFiles)
        bm.AddWorkbook "selectBook_" & i, FO.MultipleFiles(i)
        sc.AddWs bm.GetWb("selectBook_" & i).Sheets(1), "selectSheet_" & i
    Next i
    
    
    
    '========================================================================
    
    Logger.Info MODULE_NAME & " 処理の終了"

End Function

