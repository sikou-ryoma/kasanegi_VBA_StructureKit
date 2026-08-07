VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================
' DateUtility
' 日付操作用ユーティリティ
'========================================================================
Option Explicit


'----------------------------------------------------------
' 月初日を取得
'----------------------------------------------------------
Public Function FirstDay(ByVal dt As Date) As Date
    FirstDay = DateSerial(Year(dt), Month(dt), 1)
End Function

'----------------------------------------------------------
' 月末日を取得
'----------------------------------------------------------
Public Function LastDay(ByVal dt As Date) As Date
    LastDay = DateSerial(Year(dt), Month(dt) + 1, 1) - 1
End Function

'----------------------------------------------------------
' 月の日数を取得
'----------------------------------------------------------
Public Function DaysInMonth(ByVal dt As Date) As Long
    DaysInMonth = Day(LastDay(dt))
End Function


'----------------------------------------------------------
' うるう年判定
'----------------------------------------------------------
Public Function IsLeapYear(ByVal targetYear As Long) As Boolean

    IsLeapYear = _
        (targetYear Mod 400 = 0) _
        Or _
        ((targetYear Mod 4 = 0) And (targetYear Mod 100 <> 0))

End Function


'----------------------------------------------------------
' yyyyMM形式
' 例：202608
'----------------------------------------------------------
Public Function FormatYYYYMM(ByVal dt As Date) As String
    FormatYYYYMM = Format$(dt, "yyyymm")
End Function

'----------------------------------------------------------
' yyyy/MM形式
' 例：2026/08
'----------------------------------------------------------
Public Function FormatYearMonth(ByVal dt As Date) As String
    FormatYearMonth = Format$(dt, "yyyy/mm")
End Function

'----------------------------------------------------------
' yyyy/MM/dd形式
'----------------------------------------------------------
Public Function FormatDate(ByVal dt As Date) As String
    FormatDate = Format$(dt, "yyyy/mm/dd")
End Function

'----------------------------------------------------------
' 日付を加算
'----------------------------------------------------------
Public Function AddDays(ByVal dt As Date, ByVal days As Long) As Date
    AddDays = DateAdd("d", days, dt)
End Function

'----------------------------------------------------------
' 月を加算
'----------------------------------------------------------
Public Function AddMonths(ByVal dt As Date, ByVal months As Long) As Date
    AddMonths = DateAdd("m", months, dt)
End Function

'----------------------------------------------------------
' 年を加算
'----------------------------------------------------------
Public Function AddYears(ByVal dt As Date, ByVal years As Long) As Date
    AddYears = DateAdd("yyyy", years, dt)
End Function

'----------------------------------------------------------
' 曜日番号取得
' 月曜=1 ～ 日曜=7
'----------------------------------------------------------
Public Function WeekdayNumber(ByVal dt As Date) As Long
    WeekdayNumber = Weekday(dt, vbMonday)
End Function

'----------------------------------------------------------
' 曜日名取得
'----------------------------------------------------------
Public Function WeekdayNameJP(ByVal dt As Date) As String

    Static names As Variant

    If IsEmpty(names) Then
        names = Array("月", "火", "水", "木", "金", "土", "日")
    End If

    WeekdayNameJP = names(Weekday(dt, vbMonday) - 1)

End Function

