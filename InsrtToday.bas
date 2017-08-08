Attribute VB_Name = "InsrtToday"
Option Explicit

Sub FormatDate()
    Dim r As Range
    Set r = ActiveCell
    If r Is Nothing Then Exit Sub
    Dim d As String
    d = Format(Date, "Long Date")
    d = Left$(d, 5) & " " & Mid(d, 6, 2) & " " & Right$(d, 3)
    r.Value = d
End Sub
