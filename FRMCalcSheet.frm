VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMCalcSheet 
   Caption         =   "入力項目"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "FRMCalcSheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRMCalcSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub cmbCustomers_Change()
    If cmbMaker.Text <> "" Then Set_MachineName
End Sub
Private Sub cmbMaker_Change()
    If cmbCustomers.Text <> "" Then Set_MachineName
End Sub
Private Sub UserForm_Initialize()
    Set_Makers
    Set_Customers
End Sub

Private Sub Set_Makers()
    Const sql As String = "SELECT name FROM makers"
    Const DB_FILE_NAME As String = "makers.db"
    
    Dim Makers As Variant
    Makers = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)
    Dim i As Long
    For i = LBound(Makers, 2) To UBound(Makers, 2)
        cmbMaker.AddItem Makers(0, i)
    Next
End Sub
Private Sub Set_Customers()
    Const sql As String = "SELECT customer_name_ja FROM payment_conditions"
    Const DB_FILE_NAME As String = "payment_conditions.sqlite3"
    
    Dim Customers As Variant
    Customers = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)
    Dim i As Long
    For i = LBound(Customers, 2) To UBound(Customers, 2)
        cmbCustomers.AddItem Customers(0, i)
    Next
End Sub

Private Function Set_MachineName()
'顧客ごとの機械型式を表示する'
    Dim sql As String
    sql = "SELECT machine_type FROM delivered_machines WHERE customer_name = " & _
          """" & cmbCustomers.Text & """" & _
          "manufacturer_name = " & """" & cmbMaker.Text & """" & "ORDER BY machine_type ASC"
          
    Const DB_FILE_NAME As String = "winckler.sqlite3"
    
    Dim Machines As Variant
    On Error GoTo ERR
    Machines = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)
    Dim i As Long
    For i = LBound(Machines, 2) To UBound(Machines, 2)
        cmbMachineType.AddItem Machines(0, i)
    Next
    Exit Function
ERR:
Debug.Print "Set_MachineName: Maybe sql error..."
End Function
