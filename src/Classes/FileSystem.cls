VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------
' ---FileSystem---
' ファイル選択や保存関する操作を簡潔に行うためのクラスモジュールです。
'------------------------------------------------------------------------
Option Explicit


Private Const CLASS_NAME As String = "FileSystem"

Private pFileNm As Variant '---select file path
Private pMultipleFiles As Variant '---select multiple files path
Private pFolderPath As Variant '---select folder path


'---取得したファイルパスの設定・取得・解放
Public Property Let fileName(ByVal fn As Variant)
    pFileNm = fn
End Property

Public Property Get fileName() As Variant
    fileName = pFileNm
End Property

Public Sub fileName_Clear()
    pFileNm = Empty
End Sub



'---取得した複数のファイルパスの設定・取得・解放
Public Property Let MultipleFiles(ByVal fns As Variant)
    pMultipleFiles = fns
End Property

Public Property Get MultipleFiles() As Variant
    MultipleFiles = pMultipleFiles
End Property

Public Sub MultipleFiles_Clear()
    pMultipleFiles = Empty
End Sub



'---取得したフォルダパスの設定・取得
Public Property Let folderPath(ByVal fp As Variant)
    pFolderPath = fp
End Property

Public Property Get folderPath() As Variant
    folderPath = pFolderPath
End Property

Public Sub folderPath_Clear()
    pFolderPath = Empty
End Sub


'---この関数は選択したファイルのフルパスを返してそのファイルを開き、キャンセル時はFalseを返す
'---キャンセル時に付いては呼び出し元で分岐を管理する必要がある
Public Function OpenFile() As Variant
    
    Dim fileName As Variant
    'MsgBox "ブックを選択してください。", vbInformation, "OpenFile"
    fileName = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm")
    OpenFile = fileName
    If fileName = False Then
        MsgBox "キャンセルが押されたため処理を中断します。", vbExclamation, "OpenFile"
        Exit Function
    Else
        Workbooks.Open fileName
    End If
    
End Function


Public Function OpenMultipleFiles() As Variant

'---OpFileの複数選択版
'---配列を返すので呼び出し元で管理する必要がある
    Dim fileNames As Variant
    Dim i As Long
    MsgBox "ブックを選択してください。", vbInformation, "OpenMultipleFiles"
    fileNames = Application.GetOpenFilename( _
        FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm", MultiSelect:=True)
    OpenMultipleFiles = fileNames
    If VarType(fileNames) = vbBoolean Then
        MsgBox "キャンセルが押されたため処理を中断します。", vbExclamation, "OpenMultipleFiles"
        Exit Function
    Else
        For i = LBound(fileNames) To UBound(fileNames)
            Workbooks.Open fileNames(i)
        Next i
    End If

End Function


Public Function SaveFile(ByVal rcvFn As String) As Variant

    Dim fileName As Variant
    Dim FF As Long

    '---拡張子によるフィルターインデックスの設定
    If rcvFn Like "*.xlsx" Then
        FF = 1
    ElseIf rcvFn Like "*.xls" Then
        FF = 2
    ElseIf rcvFn Like "*.xlsm" Then
        FF = 3
    Else
        FF = 1
    End If

    MsgBox "保存先を選択してください。", vbInformation, "SaveFile"

    On Error GoTo ErrHandler

    fileName = Application.GetSaveAsFilename( _
        InitialFileName:=rcvFn, _
        FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm", _
        FilterIndex:=FF)

    If fileName = False Then
        MsgBox "キャンセルが押されました。", vbExclamation, "SaveFile"
        SaveFile = False
        Exit Function
    End If

    '---現在のブックを保存(必要に応じてThisWorkbookかActiveWorkbookを使い分けてください)
    ActiveWorkbook.SaveAs fileName
    SaveFile = fileName
    Exit Function

ErrHandler:
    SaveFile = False
    MsgBox "保存に失敗しました。ファイル形式やファイル名を確認してください。" & vbCrLf & _
           "エラーコード：" & Err.number & vbCrLf & "エラー内容：" & Err.Description, vbExclamation, "SaveFile"
    Err.Clear

End Function


'---この関数は選択したフォルダのパスを返す、キャンセル時はFalseを返す
'---キャンセル時に付いては呼び出し元で分岐を管理する必要がある
Public Property Get GetFolderPath() As Variant

    MsgBox "指定するフォルダを選択して下さい。", vbInformation, "GetFolderPath"
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Not .Show Then
            GetFolderPath = False
            MsgBox "キャンセルが押されました。", vbExclamation, "GetFolderPath"
            Exit Property
        End If
        GetFolderPath = .SelectedItems(1)
    End With

End Property


'---最終的な "\" を除去して返す
Public Function NormalizePath(ByVal path As String) As String

    If Right(path, 1) = "\" Then
        NormalizePath = Left(path, Len(path) - 1)
    Else
        NormalizePath = path
    End If
    
End Function


'---1つ上のフォルダのパスを取得
Public Function parentPath(ByVal fullPath As String) As String

    fullPath = NormalizePath(fullPath)
    
    Dim lastSep As Long
    lastSep = InStrRev(fullPath, "\")
    
    If lastSep > 0 Then
        parentPath = Left(fullPath, lastSep - 1)
    Else
        parentPath = fullPath ' パス区切りがない（例外）
    End If
    
End Function


'---指定された階層数だけ上のフォルダパスを返す
Public Function UpPath(ByVal fullPath As String, Optional ByVal levelUp As Long = 1) As String

    Dim i As Long
    
    fullPath = NormalizePath(fullPath)
    
    For i = 1 To levelUp
        fullPath = parentPath(fullPath)
    Next i
    
    UpPath = fullPath
End Function


Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = False
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function


Public Sub OpenExplorer(Optional ByVal targetPath As String = "")
    
    Shell "explorer.exe """ & targetPath & """", vbNormalFocus
End Sub


Function FindFolder(ByVal StartPath As String, ByVal TargetName As String) As String

    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim result As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(StartPath)

    For Each subFolder In folder.SubFolders

        If StrComp(subFolder.Name, TargetName, vbTextCompare) = 0 Then
            FindFolder = subFolder.path
            Exit Function
        End If

        result = FindFolder(subFolder.path, TargetName)

        If result <> "" Then
            FindFolder = result
            Exit Function
        End If

    Next

End Function


