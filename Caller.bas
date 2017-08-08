Attribute VB_Name = "Caller"
Option Explicit
Const Path As String = "\\LS410D760\share\【6】見積・ﾌﾟﾛｼﾞｪｸﾄ・ｵｰﾀﾞｰ番号取得表\2017 案件表.xlsm"
Const projectInventoryPath = "\\LS410D760\share\【6】見積・ﾌﾟﾛｼﾞｪｸﾄ・ｵｰﾀﾞｰ番号取得表\【2017】見積・ﾌﾟﾛｼﾞｪｸﾄ・ｵｰﾀﾞｰ番号取得表\2017  ﾌﾟﾛｼﾞｪｸﾄトNo.ｵｰﾀﾞｰNo. 取得表.xlsx"
Const QuotationNoInventoryPath = "\\LS410D760\share\【6】見積・ﾌﾟﾛｼﾞｪｸﾄ・ｵｰﾀﾞｰ番号取得表\【2017】見積・ﾌﾟﾛｼﾞｪｸﾄ・ｵｰﾀﾞｰ番号取得表\\2017  見積番号取得表 最新版.xlsx"

Public Sub ShowForm()
    Dim bk As Workbook
    Set bk = Workbooks.Open(fileName:=Path)
    '実行できない
'    Application.Run "2017 案件表.xlsm!ShowForm"
End Sub

Public Sub ShowBooks()
    Dim n As Long
    n = InputBox("開きたいブック番号を入力してください。 1= プロジェクト番号採番表　2= 見積書番号採番表")

    Select Case n
        Case 1
            ShowProjectNoInventry
        Case 2
            ShowQuotationNoInventry
        Case Else
        
    End Select

End Sub

Private Sub ShowProjectNoInventry()
    Dim bk As Workbook
    Set bk = Workbooks.Open(fileName:=projectInventoryPath)
End Sub
Private Sub ShowQuotationNoInventry()
    Dim bk As Workbook
    Set bk = Workbooks.Open(fileName:=QuotationNoInventoryPath)
End Sub

Public Sub ShowCalcForm()
    FRMCalcSheet.Show False
End Sub
