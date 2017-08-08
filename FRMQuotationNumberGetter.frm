VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMQuotationNumberGetter 
   Caption         =   "���͍���"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   OleObjectBlob   =   "FRMQuotationNumberGetter.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FRMQuotationNumberGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DB_FILE_NAME As String = "winckler.sqlite3"

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnEnter_Click()
InsertData
End Sub

Private Sub cmb�q��_Change()
    'ResetForm
    Set_CustomerFactory
    If cmb���[�J�[.Text <> "" Then Set_MachineName
End Sub
Private Sub cmb���[�J�[_Change()
    'ResetForm
    If cmb�q��.Text <> "" Then Set_MachineName
End Sub
Private Sub cmb�Ώۋ@��_Change()
    Set_�V���A��No
End Sub

Private Sub UserForm_Initialize()
    Set_Date
    Set_Makers
    Set_Customers
End Sub

Private Sub Set_Makers()
    Const sql As String = "SELECT DISTINCT manufacturer_name FROM delivered_machines ORDER BY manufacturer_name"
    
    Dim Makers As Variant
    Makers = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)
    cmb���[�J�[.List = Array()
    Dim i As Long
    For i = LBound(Makers, 2) To UBound(Makers, 2)
        cmb���[�J�[.AddItem Makers(0, i)
    Next
End Sub
Private Sub Set_Customers()
    Const sql As String = "SELECT DISTINCT customer_name FROM delivered_machines ORDER BY customer_name ASC"
    
    Dim Customers As Variant
    Customers = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)
    cmb�q��.List = Array()
    Dim i As Long
    For i = LBound(Customers, 2) To UBound(Customers, 2)
        cmb�q��.AddItem Customers(0, i)
    Next
End Sub
Private Sub Set_CustomerFactory()
    Dim sql As String
    sql = "SELECT DISTINCT customer_factory FROM delivered_machines " & _
                            " WHERE customer_name=" & """" & cmb�q��.Text & """" & _
                            " ORDER BY customer_factory ASC"
    
    Dim Factories As Variant
    Factories = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)
    cmb�q��H��.List = Array()
    Dim i As Long
    For i = LBound(Factories, 2) To UBound(Factories, 2)
        cmb�q��H��.AddItem Factories(0, i)
    Next
End Sub

Private Function Set_MachineName()
    Dim sql As String
    sql = "SELECT machine_type FROM delivered_machines WHERE customer_name = " & _
          """" & cmb�q��.Text & """" & " AND " & _
          " manufacturer_name = " & """" & cmb���[�J�[.Text & """" & _
          " ORDER BY machine_type ASC"
    Debug.Print sql
    
    On Error GoTo ERR
    Dim Machines As Variant
    Machines = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)

    If Util.IsArrayEx(Machines) = 0 Then
        cmb�Ώۋ@��.List = Array()
        cmb�Ώۋ@��.Enabled = False
        Exit Function
    Else
        cmb�Ώۋ@��.Enabled = True
    End If
    
    Dim i As Long
    For i = LBound(Machines, 2) To UBound(Machines, 2)
        cmb�Ώۋ@��.AddItem Machines(0, i)
    Next
    Exit Function
ERR:
Debug.Print "Set_MachineName: error..."
End Function
Private Function Set_�V���A��No()
    Dim sql As String
    sql = "SELECT maker_order_id FROM delivered_machines WHERE customer_name = " & _
          """" & cmb�q��.Text & """" & " AND " & _
          " manufacturer_name=" & """" & cmb���[�J�[.Text & """" & _
        " AND " & " machine_type=" & """" & cmb�Ώۋ@��.Text & """" & _
          " ORDER BY maker_order_id ASC"
    Debug.Print "sql: " & sql
    
    On Error GoTo ERR
    Dim SerialNos As Variant
    SerialNos = sqlite_no_ADODB.SearchAll(sql, DB_FILE_NAME)
    
    If Util.IsArrayEx(SerialNos) = 0 Then
        cmb�V���A��No.List = Array()
        cmb�V���A��No.Enabled = False
        Exit Function
    Else
        cmb�V���A��No.Enabled = True
    End If
    cmb�V���A��No.List = Array()
    Dim i As Long
    For i = LBound(SerialNos, 2) To UBound(SerialNos, 2)
        cmb�V���A��No.AddItem SerialNos(0, i)
    Next
    Exit Function
ERR:
Debug.Print "Set_�V���A��No: error..."
End Function

Private Sub ResetForm()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Debug.Print TypeName(ctrl)
        Select Case TypeName(ctrl)
            Case "ComboBox"
                ctrl.List = Array()
            Case "Textbox"
                ctrl.Text = ""
        End Select
    Next ctrl
End Sub
Private Sub Set_Date()
    txt���ϓ�.Text = Format(Date, "long date")
    txt���ϓ�.Enabled = False
End Sub
    
Sub InsertData()
'15�񂠂邯��13���ڂ����Ȃ�����G���[�ɂȂ��Ă�
    Const TBL_NAME As String = "winckler_quotation_overview"
    Dim sql As String
    sql = "INSERT INTO " & TBL_NAME & "(customer_person, customer_name, machine_type," & _
    "maker_order_id, created_at, quotation_content, customer_factory, maker, created_by," & _
    "purchased_from_foreign, maker_quotation_id, purchased_from_domestic," & _
    "winckler_quotation_id, is_main_machine, is_over_tenmil) VALUES(" & _
    """" & cmb�q��S��.Text & """" & "," & _
    """" & cmb�q��.Text & """" & "," & _
    """" & cmb�Ώۋ@��.Text & """" & "," & _
    """" & cmb�V���A��No.Text & """" & "," & _
    """" & txt���ϓ�.Text & """" & "," & _
    """" & txt���ϓ��e.Text & """" & "," & _
    """" & cmb�q��H��.Text & """" & "," & _
    """" & cmb���[�J�[.Text & """" & "," & _
    """" & txt�쐬��.Text & """" & "," & _
    """" & txt�d����.Text & """" & "," & _
    """" & txtQuotationno.Text & """" & "," & _
    """" & txt�����d����.Text & """" & "," & _
    """" & txt���ϔԍ�.Text & """" & _
    """" & chk�{�@.Value & """" & _
    """" & ������Over10mil.Value & """" & _
    ")"
    Debug.Print sql
        sqlite_no_ADODB.Insert sql, DB_FILE_NAME

End Sub
