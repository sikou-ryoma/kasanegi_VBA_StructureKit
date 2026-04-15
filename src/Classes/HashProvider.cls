VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function DJB2(ByVal text As String) As Long
#If Win64 Then
    Dim hash As LongLong
#Else
    Dim hash As Long
#End If

    Dim i As Long
    hash = 5381

    For i = 1 To Len(text)
#If Win64 Then
        hash = ((hash * 33) + Asc(Mid$(text, i, 1))) And &H7FFFFFFF
#Else
        hash = (hash + ((hash And &H3FFFFFF) * 32)) And &H7FFFFFFF
        hash = (hash + Asc(Mid$(text, i, 1))) And &H7FFFFFFF
#End If
    Next i

    DJB2 = CLng(hash)
End Function
