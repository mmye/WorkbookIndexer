Attribute VB_Name = "Caller"
Option Explicit
Const Path As String = "\\LS410D760\share\�y6�z���ρE��ۼު�āE���ް�ԍ��擾�\\2017 �Č��\.xlsm"
Const projectInventoryPath = "\\LS410D760\share\�y6�z���ρE��ۼު�āE���ް�ԍ��擾�\\�y2017�z���ρE��ۼު�āE���ް�ԍ��擾�\\2017  ��ۼު�ăgNo.���ްNo. �擾�\.xlsx"
Const QuotationNoInventoryPath = "\\LS410D760\share\�y6�z���ρE��ۼު�āE���ް�ԍ��擾�\\�y2017�z���ρE��ۼު�āE���ް�ԍ��擾�\\\2017  ���ϔԍ��擾�\ �ŐV��.xlsx"

Public Sub ShowForm()
    Dim bk As Workbook
    Set bk = Workbooks.Open(fileName:=Path)
    '���s�ł��Ȃ�
'    Application.Run "2017 �Č��\.xlsm!ShowForm"
End Sub

Public Sub ShowBooks()
    Dim n As Long
    n = InputBox("�J�������u�b�N�ԍ�����͂��Ă��������B 1= �v���W�F�N�g�ԍ��̔ԕ\�@2= ���Ϗ��ԍ��̔ԕ\")

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
