Attribute VB_Name = "AutoCellLabeler"
Option Explicit

Const str本機 As String = "本機"
Const str1000万以上 As String = "1000万" & vbLf & "以上"
Const str見積番号 As String = "見積番号"
Const str客先 As String = "客先"
Const str客先担当 As String = "客先担当"
Const str客先工場 As String = "客先工場"
Const str対象機器 As String = "対象機器"
Const str機械番号 As String = "シリアルNo."
Const str見積内容 As String = "見積内容"
Const str見積日 As String = "見積日"
Const str作成者 As String = "作成者"
Const str担当者 As String = "担当者"
Const strQuotationNo As String = "Quotation no."
Const str追加 As String = "追加"
Const strプロジェクトナンバー As String = "プロジェクトナンバー"
Const strオーダーナンバー As String = "オーダーナンバー"
Const str入力日 As String = "入力日(YY/MM/DD)"
Const strプロジェクト内容 As String = "プロジェクト内容"
Const str受注見積番号 As String = "受注見積番号"
Const str仕入先 As String = "国内仕入先"
Const str備考 As String = "備考"
Const str部品 As String = "部品"
Const str作業日 As String = "作業日"
Const HeaderRow As Long = 2

Sub LoopSheets()
    Dim st As Worksheet
    For Each st In ActiveWorkbook.Sheets
        LabelColumnn st
    Next st
End Sub

Sub LabelColumnn(st As Worksheet)
    Const HeaderRow As Long = 2
    Dim c As Long
    Dim LastCol As Long
    LastCol = Cells(2, st.Columns.Count).End(xlToLeft).Column
    For c = 1 To LastCol
        Dim r As Range
        Set r = st.Cells(HeaderRow, c)

        Dim Prefix As String
'        If st.Name Like "[\(\)-&（）]" Then
            Prefix = RegularExpressions.RegexReplace(st.name, "\(|\)", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "（", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "）", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "&", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "-", "")
 '       End If
        Select Case r.Value
            Case str本機
                r.name = Prefix & "_" & str本機
            Case str1000万以上
                r.name = Prefix & "_" & str1000万以上
            Case str見積番号
                r.name = Prefix & "_" & str見積番号
            Case str客先
                r.name = Prefix & "_" & str客先
            Case strQuotationNo
                r.name = Prefix & "_" & strQuotationNo
            Case str追加
                r.name = Prefix & "_" & str追加
            Case strプロジェクトナンバー
                r.name = Prefix & "_" & strプロジェクトナンバー
            Case strオーダーナンバー
                r.name = Prefix & "_" & strオーダーナンバー
            Case str入力日
                r.name = Prefix & "_" & str入力日
            Case str客先工場
                r.name = Prefix & "_" & str客先工場
            Case str対象機器
                r.name = Prefix & "_" & str対象機器
            Case str機械番号
                r.name = Prefix & "_" & str機械番号
            Case strプロジェクト内容
                r.name = Prefix & "_" & strプロジェクト内容
            Case str受注見積番号
                r.name = Prefix & "_" & str受注見積番号
            Case str作成者
                r.name = Prefix & "_" & str作成者
            Case str担当者
                r.name = Prefix & "_" & str担当者
            Case str仕入先
                r.name = Prefix & "_" & str仕入先
            Case str備考
                r.name = Prefix & "_" & str備考
            Case str部品
                r.name = Prefix & "_" & str部品
            Case str作業日
                r.name = Prefix & "_" & str作業日
        End Select
    Next
End Sub





