VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM�^�[�~�i�� 
   Caption         =   "���Ϗ����J��..."
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   OleObjectBlob   =   "FRM�^�[�~�i��.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FRM�^�[�~�i��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NewStName As String
Dim mLists As Collection
Dim IndexFile As String
Const UboundArr As Long = 6
Const IndexPath As String = "\\LS410D760\share\������\�s1�t���ρE����\1. ���Ϗ�" '�����̃Z�p���[�^�[�͂Ȃ��Ŏw�肷��
Const strInstruction As String = "���Ϗ��ԍ�����͂���Enter�������Ă�������"

Enum WriteReadMode
    Writable = True
    ReadOnly = False
End Enum

Enum RevisionStatus
    FirstRevision = 0
    HasRevisedBefore = 1
End Enum

Private Sub txtTerminal_Change()
    Dim buf As String
    buf = txtTerminal.Text
    txtTerminal.Text = UCase(buf)
    If txtStatus.Caption = "" Then txtStatus.Caption = strInstruction
    If txtStatus.Caption <> "" Then
        txtStatus.Caption = Empty
        txtStatus.Caption = strInstruction
    End If
End Sub

Private Sub UserForm_Click()
    txtTerminal.SetFocus
End Sub

Sub LoadWks()
    Set mLists = LoadIndexLists
    If mLists Is Nothing Then
        MsgBox "�C���f�b�N�X����ł�"
        Exit Sub
    End If
End Sub

Private Sub txtTerminal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print "KEY CODE:" & KeyCode
    Select Case KeyCode
        Case 27 ' ESC
            Me.Hide
            Exit Sub
        Case 13
            If LenB(txtTerminal.Text) = 0 Then Exit Sub
            Dim Key As String
            Key = StrConv(txtTerminal.Text, vbNarrow)
            Key = Trim$(Key)
            Dim st As Worksheet
            Select Case Right$(Key, 2)
                Case "-R"
                    Key = Left$(Key, Len(Key) - 3)
                    Dim IsRevision As Boolean
                    IsRevision = True
                    Key = UCase(Key)
                    '���r�W�����쐬
                    Set st = OpenSheet(Key, IsRevision, Writable)
                Case "-W"
                    'Public scope: Dim Writable As Boolean
                    Key = Left$(Key, Len(Key) - 3)
                    Key = UCase(Key)
                    '�������݃��[�h�łЂ炭
                    Set st = OpenSheet(Key, , Writable)
                Case Else
                    '�ǂݎ���p�łЂ炭
                    Set st = OpenSheet(Key, , ReadOnly)
                End Select
                
        '���r�W�����V�[�g���̓��t�Ƃ����X�V����
        If Not st Is Nothing Then UpdateContents st, NewStName
        Set st = Nothing
    End Select
    
End Sub

Sub UpdateContents(st, name)

    Dim rQuotationNumber As Range
    Set rQuotationNumber = st.Cells.Find(What:="No.", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
            , MatchByte:=False, SearchFormat:=False)
    rQuotationNumber.Value = "No." & name

    Dim rDate As Range
    Set rDate = st.Cells.Find(What:="Date:", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
            , MatchByte:=False, SearchFormat:=False)
    Dim i As Long
    Dim r As Range
    For i = 0 To 10
        Set r = rDate.offset(0, i)
        If r.Value Like "*�N*��*��" Then r.Value = GetToday
    Next
    Set rQuotationNumber = Nothing
End Sub
Private Function GetToday() As String
    Dim d As String
    d = Format(Date, "Long Date")
    d = Left$(d, 5) & " " & Mid(d, 6, 2) & " " & Right$(d, 3)
    GetToday = d
End Function

Function OpenSheet(Key As String, Optional IsRevision As Boolean = False, _
                    Optional Writable As Boolean = False) As Worksheet
    On Error Resume Next

    Dim v As Variant
    v = Split(mLists(Key), vbTab)
    If IsEmpty(v) Then
        txtStatus.Caption = Empty
        txtStatus.Caption = "���Ϗ���������܂���B"
        Exit Function
    End If

    Dim Path As String
    Path = v(6) & v(5)
    Dim stName As String
    stName = v(7)
    Dim File As String
    File = Dir(Path)

    Select Case File
        Case ""
            '���Ϗ��u�b�N���C���f�b�N�X�Ɍ�����Ȃ������ꍇ
            '�������Ȃ�
        Case Else
            On Error Resume Next
            Select Case Writable
                Dim bk As Workbook
                Case True
                    Set bk = Workbooks.Open(fileName:=Path, ReadOnly:=False)
                Case False
                    Set bk = Workbooks.Open(fileName:=Path, ReadOnly:=True)
                    bk.Select
            End Select

            If IsRevision Then
                If stName <> "" Then
                    Dim st As Worksheet
                    Application.DisplayAlerts = False
                    bk.Sheets(stName).Copy After:=bk.Sheets(stName)
                    Set st = ActiveSheet
                    Application.DisplayAlerts = False

                    '���r�W�����̃V�[�g��������
                    'Public scope: NewStName'
                    st.name = GetNewRevisionNumber.GetNumber(bk, stName)
                    NewStName = st.name
                    st.Activate
                    Range("a1").Select
                    Set OpenSheet = st
                End If
            Else
                If stName <> "" Then
                    bk.Sheets(stName).Activate
                    Range("a1").Select
                    If Not OpenSheet Is Nothing Then Set OpenSheet = st
                End If
                '�t�H�[�������
                Me.Hide
            End If
    End Select
    On Error GoTo 0
    Set bk = Nothing
End Function

Private Sub UserForm_Initialize()
Attribute UserForm_Initialize.VB_ProcData.VB_Invoke_Func = "H\n14"
    LoadWks
    txtStatus.Caption = strInstruction
End Sub

Private Function LoadIndexLists() As Collection
    Dim v As Variant
    Dim a As String, l As String, z As String
    Dim Lists As New Collection
    Dim List As Variant

    IndexFile = IndexPath & "\" & "index.txt"
    If Dir(IndexFile) = "" Then MsgBox "�C���f�b�N�X�t�@�C����������܂���B", vbCritical

    On Error Resume Next
    Open IndexFile For Input As #1
    Do Until EOF(1)
        ERR.Number = 0
        Line Input #1, a
        List = Split(a, vbTab)
        Dim Key As String, Content As String
        Key = List(0)
        Content = GetContent(List) 'Content ���w���BVariant�z���Ԃ��ׂ����H�H
        On Error Resume Next
        Lists.Add Content, Key
    Loop
    On Error GoTo 0
    Close #1
    Set LoadIndexLists = Lists
    Exit Function
ERR:
Close #1

End Function

Private Function GetContent(List As Variant) As String
    Dim v As Variant
    Dim buf As String
    For Each v In List
        buf = buf & vbTab & v
    Next v
    GetContent = buf
End Function

Private Sub UserForm_Terminate()
    Set mLists = Nothing
End Sub
