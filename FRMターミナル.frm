VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMターミナル 
   Caption         =   "見積書を開く..."
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   OleObjectBlob   =   "FRMターミナル.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRMターミナル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NewStName As String
Dim mLists As Collection
Dim IndexFile As String
Const UboundArr As Long = 6
Const IndexPath As String = "\\LS410D760\share\◆事務\《1》見積・注文\1. 見積書" '末尾のセパレーターはなしで指定する
Const strInstruction As String = "見積書番号を入力してEnterを押してください"

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
        MsgBox "インデックスが空です"
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
                    'リビジョン作成
                    Set st = OpenSheet(Key, IsRevision, Writable)
                Case "-W"
                    'Public scope: Dim Writable As Boolean
                    Key = Left$(Key, Len(Key) - 3)
                    Key = UCase(Key)
                    '書き込みモードでひらく
                    Set st = OpenSheet(Key, , Writable)
                Case Else
                    '読み取り専用でひらく
                    Set st = OpenSheet(Key, , ReadOnly)
                End Select
                
        'リビジョンシート内の日付とかを更新する
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
        If r.Value Like "*年*月*日" Then r.Value = GetToday
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
        txtStatus.Caption = "見積書が見つかりません。"
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
            '見積書ブックがインデックスに見つからなかった場合
            '何もしない
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

                    'リビジョンのシート名をつける
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
                'フォームを閉じる
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
    If Dir(IndexFile) = "" Then MsgBox "インデックスファイルが見つかりません。", vbCritical

    On Error Resume Next
    Open IndexFile For Input As #1
    Do Until EOF(1)
        ERR.Number = 0
        Line Input #1, a
        List = Split(a, vbTab)
        Dim Key As String, Content As String
        Key = List(0)
        Content = GetContent(List) 'Content がヘン。Variant配列を返すべきか？？
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
