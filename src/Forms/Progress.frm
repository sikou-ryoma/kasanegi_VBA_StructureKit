VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progress 
   Caption         =   "UserForm1"
   ClientHeight    =   1365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   OleObjectBlob   =   "Progress.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isCancelPb As Boolean '中断時にTrueにする

Private pProgressBar As MSForms.Label 'ラベル：動的追加
Private pMaxValue As Long 'プログレスバー最大値
Private pBarColor As Long 'プログレスバー色
Private pCurValue As Double 'プログレスバー現在値
Private pInteractive As Long '割込み

'最大値プロパティ
Public Property Let MaxValue(aMaxValue As Long)
    pMaxValue = aMaxValue
End Property
Public Property Get MaxValue() As Long
    MaxValue = pMaxValue
End Property

'プログレスバー色プロパティ
Public Property Let BarColor(aBarColor As Long)
    pBarColor = aBarColor
End Property

'割込み拒否プロパティ
Public Property Let Interactive(aInteractive As Boolean)
    pInteractive = aInteractive
End Property

'フォーム表示入り口
Public Sub ShowModeless(Optional ByVal strTitle As String = "")

    'ラベルコントロール追加
    Set pProgressBar = Me.FrameProgress.Controls.Add("Forms.Label.1", "lblProgress")
    If pBarColor = 0 Then pBarColor = RGB(0, 0, 128)
    pProgressBar.Width = 0
    pProgressBar.Height = Me.FrameProgress.Height
    pProgressBar.BackColor = pBarColor
  
    'プログレスバーの背景をへこませる
    Me.FrameProgress.SpecialEffect = fmSpecialEffectSunken
  
    '割込み拒否の設定
    If pInteractive = False Then
        Me.Enabled = False '好みで設定
        Application.Interactive = False
        Application.EnableCancelKey = xlDisabled
    End If
  
    'フォームをモードレスで表示
    Me.Caption = ""
    Me.Show vbModeless
End Sub

'プログレス進捗：指定値
Public Sub Value(ByVal aValue As Double, Optional ByVal strTitle As String = "")
    'プログレスバー値変更
    pCurValue = aValue
  
    '最大値判定
    If pCurValue > pMaxValue Then
        pCurValue = pMaxValue
    End If
  
    'プログレスバーの描画
    pProgressBar.Width = pCurValue * Me.FrameProgress.Width / pMaxValue
    If Me.Caption <> strTitle Then
        Me.Caption = strTitle
    End If
  
    '再描画
    'Me.Repaint 'これだと「応答なし」が出てしまう
    DoEvents
End Sub

'プログレス進捗：加算
Public Sub ValueAdd(ByVal aValue As Double, Optional ByVal strTitle As String = "")
    pCurValue = pCurValue + aValue
    Call Value(pCurValue, strTitle)
End Sub

'フォーム終了
Public Sub SelfClose()
    Unload Me
End Sub

'正規終了以外をキャンセル
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If pInteractive Then
            If MsgBox("処理を中断しますか?", vbYesNo, "中断確認") = vbYes Then
                isCancelPb = True
            Else
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End If
End Sub

'フォーム終了時に割込み拒否を戻す
Private Sub UserForm_Terminate()
    If pInteractive = False Then
        Application.Interactive = True
        Application.EnableCancelKey = xlInterrupt
    End If
End Sub

Public Sub Cpt(ByVal txt As String)
    Progress.Caption = txt
End Sub
