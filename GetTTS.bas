Attribute VB_Name = "GetTTS"
Option Explicit
Dim tmpSt As Worksheet

Sub ShowTTS()
    Application.ScreenUpdating = False
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.application")
    IE.Visible = False
    IE.navigate "http://www.bk.mufg.jp/gdocs/rate/real_01.html"
    Do While IE.busy Or IE.readystate < 4 ' READYSTATE_COMPLETEの値
        DoEvents
    Loop
    
    Dim CurrentSt As Worksheet
    Set CurrentSt = ActiveSheet
    
    Call MakeList(IE)
    Dim Rate As String
    Rate = GetTTS
    Call DeleteSheet '削除
    IE.Quit 'IEを閉じる（片付け）
    Application.ScreenUpdating = True
       
    CurrentSt.Select
    Set CurrentSt = Nothing
    MsgBox "現在のTTSレート：" & Rate
    
End Sub
Private Sub MakeList(objIE As Object)

'一時シートにウェブページの内容を書き込みます
    Dim n As Long
    Dim r As Long
    Dim Doc As Object
    Dim tmpSt As Worksheet
    Dim objTD As Object
    Dim objTag As Object
    n = 0
    r = 0
    
    Set tmpSt = ThisWorkbook.Sheets.Add
    With tmpSt
        With .Range("a1", ActiveCell.SpecialCells(xlLastCell))
            .ClearContents
            .NumberFormatLocal = "G/標準"
        End With
        
        Set Doc = objIE.document
        Set objTD = Doc.getElementsByTagName("TD")
        For Each objTag In objTD
            r = r + 1
            .Cells(r + 1, 1) = objTag.tagName
            .Cells(r + 1, 2) = n
            .Cells(r + 1, 3) = r
            .Cells(r + 1, 4) = objTag.innerText
        Next objTag
    End With
End Sub

Private Function GetTTS() As String
'一時シートからTTSの情報を取り出します
    With tmpSt
        Dim i As Long
        For i = 1 To Range("A10000").End(xlUp).row
            If Range("D" & i) = "EUR (ユーロ)" Then
                GetTTS = Range("D" & i).offset(1, 0).Value
                Exit Function '必要なネタを取ったら出ます
            End If
        Next
    End With
End Function

Private Sub DeleteSheet()
'一時シートを削除します
    Application.DisplayAlerts = False '「本当に削除しますか？」を表示させない
    On Error Resume Next
    tmpSt.Delete
    Application.DisplayAlerts = True '元に戻す
    On Error GoTo 0
End Sub
