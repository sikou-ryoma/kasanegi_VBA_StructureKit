Attribute VB_Name = "ProgressManager"
Option Explicit

Public progFrm As ProgressForm

Public Sub ShowProgress(Optional ByVal seconds As Long = 2)
    
    Set progFrm = New ProgressForm
    Dim waitUntil As Date
    waitUntil = Now + TimeSerial(0, 0, seconds)

    With progFrm
        .StartUpPosition = 0
        .Left = 150
        .Top = 120
        .MaxValue = 100
        .BarColor = RGB(0, 0, 128)
        .Interactive = True
        .ShowModeless "ŠJŽn‚µ‚Ü‚·"
    End With
    
    Application.Wait waitUntil
    
End Sub

Public Function IsFormOpen(ByVal FormName As String) As Boolean

    Dim frm As Object

    For Each frm In VBA.UserForms
        If frm.Name = FormName Then
            IsFormOpen = True
            Exit Function
        End If
    Next frm

End Function

