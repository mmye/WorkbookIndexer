VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirLooper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' # �󂯎�������[�g�f�B���N�g�����ċA���[�v����N���X
' # ���[�v���ł����Ȃ������͊֐����v���C�x�[�g�T�u���[�`���ł�邱��

' # ��邱�ƁF���[�N�u�b�N�t�@�C����Sha1�n�b�V�����v�Z���āA�X�V���������Ƃ��̂ݒ��g���X�L��������
' #             �X�L�����������e�͔z��Ƃ��ĕԂ�l�Ŗ߂��B��������Ƃ�SQLite DB�ɏ������ށB
' # �e�[�u���͂ӂ��g���B�ЂƂ̓n�b�V���ƃu�b�N���̃y�A�l�A�ӂ��̓V�[�g���ƃu�b�N�̃t���p�X�̃y�A�l

Dim Path As String
Dim Books As Collection

Sub Init(p)
    If p <> "" Then
        If Right$(p, 1) <> "$" Then p = p & "\"
        Path = p
    End If
End Sub

Function Indexing() As Collection
'�C���f�b�N�X���R���N�V�����ŕԂ�
    Set Books = New Collection
    Dim ret
    ret = ScanWorkbook(Path)
    Dim v
    Stop
    'v = IsArrayEx(ret)
    'If v = 1 Then updateDB (ret)
End Function

Private Function ScanWorkbook(p)
    Dim File As String
    File = Dir(Path & Application.PathSeparator & "*.xls*")
    
    Do While File <> ""
        If Not Left$(File, 1) = "~" Then '�o�b�N�A�b�v�t�@�C����B���t�@�C���݂����Ȃ��͔̂�����
            Dim wkb As Workbook
            Dim wkbpath
            wkbpath = Path & Application.PathSeparator & File
            
            '�n�b�V�����r����͈̂�x�X�L�����������Ƃ�����u�b�N�̂�
            '�u�b�N�̃n�b�V���l�͓Ɨ������e�[�u���ɕۑ�����
            Dim sha1
            sha1 = hash.calcSha1(wkbpath)
            Debug.Print sha1
            Set wkb = EvalBookIsOpen(wkb, wkbpath) '���ꂩ���J���Ă����炷������
            ActiveWindow.Visible = False
            If Not wkb Is Nothing Then
                Debug.Print wkb.name
                Scan wkb, wkbpath
                CloseWkb wkb
            End If
        End If
        File = Dir()
    Loop
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            ScanWorkbook (f.Path)
        Next
    End With
    Stop
    ScanWorkbook = Books

    Set wkb = Nothing
    Set f = Nothing
End Function
Private Sub Scan(wkb, wkbpath)
    Dim IDs
    IDs = GetSheetNames(wkb)
    On Error Resume Next '�����L�[���o�Ă��邯�ǁA����͊֌W�Ȃ��V�[�g������X�L�b�v����OK
    Dim v
    For Each v In IDs
        Books.Add Item:=wkbpath, Key:=v
    Next
    On Error GoTo 0
End Sub
Private Sub CloseWkb(wkb As Workbook)
    wkb.Saved = True
    wkb.Close
End Sub
Private Function GetSheetNames(wkb)
    Dim names
    ReDim names(wkb.Sheets.Count - 1)
    Dim st
    Dim c: c = 0
    For Each st In wkb.Sheets
        names(c) = st.name
        Debug.Print names(c)
        c = c + 1
    Next
    GetSheetNames = names
End Function

Private Sub updateDB(ret)

End Sub
Private Function EvalBookIsOpen(wkb, Path) As Workbook
'�g�p���̃t�@�C�����J���Ȃ�
    Application.DisplayAlerts = False
    Set wkb = Workbooks.Open(fileName:=Path, Notify:=False)
    If wkb.ReadOnly Then
        wkb.Close
    End If
    Set EvalBookIsOpen = wkb
    Application.DisplayAlerts = True
End Function