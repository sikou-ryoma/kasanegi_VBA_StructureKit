VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IptPass 
   Caption         =   "パスワードの入力"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   OleObjectBlob   =   "IptPass.frx":0000
End
Attribute VB_Name = "IptPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'---テキストボックスの内容をTagに格納し呼び出し元でインプットキーとして扱う
'---パスワードの設定はZ_ModConfig内へ記載
'---具体的な呼び出し例はZ_OpenFiles内に有り


Private Sub Cancel_Btn_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub OK_Btn_Click()
    Me.Tag = Me.Pass_txt.value
    Me.Hide
End Sub


'---キー入力からエンターでOKボタンにアクセス
Private Sub Pass_txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Call OK_Btn_Click
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = ThisWorkbook.Windows(1).Left + 80
    Me.Top = ThisWorkbook.Windows(1).Top + 110
    Me.Pass_txt.MaxLength = 16
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Tag = ""
        Me.Hide
    End If
End Sub
