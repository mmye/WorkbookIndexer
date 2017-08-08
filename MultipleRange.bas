Attribute VB_Name = "MultipleRange"
Option Explicit

Sub SelectMultipleRanges()

    Dim col As Long
    col = Range("’P‰¿").Column
    Dim i As Long
    For i = 39 To 50
        Dim r As Range
        If Cells(i, col) <> "" And IsNumeric(Cells(i, col)) Then
            If r Is Nothing Then
                Set r = Cells(i, col)
            Else
                Set r = Union(r, Cells(i, col))
            End If
        End If
    Next
    r.Select
    
    Dim p As String
    p = InputBox("input prices")

    Dim v As Variant
    v = Split(p, ",")
    
    Dim cell As Range
    Dim c As Long: c = 0
    On Error GoTo ERR
    For Each cell In r
        cell.Value = v(c)
        c = c + 1
    Next
    Exit Sub
ERR:
If ERR.Number = 9 Then MsgBox "“ü—Í€–Ú‚ª‘«‚è‚Ä‚¢‚Ü‚¹‚ñB"
End Sub
