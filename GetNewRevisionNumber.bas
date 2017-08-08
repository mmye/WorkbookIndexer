Attribute VB_Name = "GetNewRevisionNumber"
Option Explicit
Sub test()
Dim myNewnum As Long
    myNewnum = GetNumber(ActiveWorkbook, "N17GRO-044")
    MsgBox myNewnum
End Sub

Function GetNumber(Book As Workbook, myStName) As String
    '引数１：操作するワークブック、引数2：リビジョン対象のシート名（リビジョン文字列を除く）
    If Right$(myStName, 2) Like "R[0-9]" Then
        myStName = Left(myStName, Len(myStName) - 2)
    End If

    Dim st As Worksheet
    Dim c As Long: c = 0
    Dim v() As String
    For Each st In Book.Sheets
        If st.name Like myStName & "R[0-9]" Then
            Dim num
            num = Right$(st.name, 1)
            ReDim Preserve v(c) As String
            v(c) = num
            c = c + 1
        End If
    Next st

    Dim maxRevision As Long
    If v(0) <> "" Then
        GetNumber = myStName & "R" & (getMax(v) + 1)
    Else
        '改定がない
        GetNumber = myStName & "R" & 1
    End If

End Function

Private Function getMax(v) As Long
    Dim i As Long
    Dim max As Long: max = 1
    For i = LBound(v) To UBound(v)
        If v(i) > max Then max = v(i)
    Next i
    getMax = max
End Function
