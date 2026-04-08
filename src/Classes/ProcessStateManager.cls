VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessStateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' ---ProcessStateManager---
' ブックを複数回処理する場合に状態を管理するクラスです。
' インスタンス作成後任意のブックをLoadすることで対象ブックに状態を示す名前空間と値を追加します。
'------------------------------------------------------------------------

Option Explicit


' 定数
'------------------------------------------------------------------------
Private Const CLASS_NAME As String = "[ProcessStateManager]"
Private Const STATE_NAME As String = "_sys_ProcessState"


' 列挙体（状態）
'------------------------------------------------------------------------
Public Enum ProcessState
    PROCESS_INIT
    WAIT_USER
    PROCESS_RESUME
    PROCESS_DONE
End Enum


' メンバ変数
'------------------------------------------------------------------------
Private m_state As ProcessState
Private m_wb As Workbook




'---初期化 or 読み込み（Name版）
Public Sub Load(ByRef wb As Workbook)

    Set m_wb = wb
    
    Dim nm As Name
    On Error Resume Next
    Set nm = m_wb.Names(STATE_NAME)
    On Error GoTo 0

    If nm Is Nothing Then
        '---無ければ新規作成
        m_state = PROCESS_INIT
        m_wb.Names.Add Name:=STATE_NAME, RefersTo:="=" & m_state
        
    Else
        '---既存状態の読み込み
        Dim val As Variant
        val = Evaluate(nm.RefersTo)
        
        If IsNumeric(val) Then
            m_state = val
        Else
            m_state = PROCESS_INIT
            nm.RefersTo = "=" & m_state
        End If
    End If

End Sub


'---状態取得
Public Property Get state() As ProcessState
    state = m_state
End Property


'---状態更新
Public Property Let state(ByVal value As ProcessState)

    If m_wb Is Nothing Then
        Err.Raise vbObjectError + 1000, , "State未初期化"
    End If
    
    '---状態遷移チェック
    If Not CanTransit(value) Then
        Err.Raise vbObjectError + 1001, , "不正な状態遷移"
    End If

    Dim nm As Name
    Set nm = m_wb.Names(STATE_NAME)

    '---永続化（式として書く）
    nm.RefersTo = "=" & value
    
    m_state = value

End Property


'---設定された状態を初期化・再設定
Public Sub Reset()

    If m_wb Is Nothing Then
        Err.Raise vbObjectError + 1005, , "State未初期化"
    End If

    Dim nm As Name
    Set nm = m_wb.Names(STATE_NAME)

    m_state = PROCESS_INIT
    nm.RefersTo = "=" & m_state

End Sub

'---状態遷移のルール
Private Function CanTransit(ByVal nextState As ProcessState) As Boolean

    Select Case m_state
        Case PROCESS_INIT
            CanTransit = (nextState = WAIT_USER)
        
        Case WAIT_USER
            CanTransit = (nextState = PROCESS_RESUME)
        
        Case PROCESS_RESUME
            CanTransit = (nextState = PROCESS_DONE)
        
        Case Else
            CanTransit = False
    End Select

End Function

