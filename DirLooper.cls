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

' # 受け取ったルートディレクトリを再帰ループするクラス
' # ループ内でおこなう処理は関数かプライベートサブルーチンでやること
' # やること：ワークブックファイルのSha1ハッシュを計算して、更新があったときのみ中身をスキャンする
' #           スキャンした内容は配列として返り値で戻す。これをあとでSQLite DBに書き込む。
' # テーブルはふたつ使う。ひとつはハッシュとブック名のペア値、ふたつはシート名とブックのフルパスのペア値

Dim Path As String

Sub Init(p)
    If p <> "" Then
        If Right$(p, 1) <> "$" Then p = p & "\"
        Path = p
    End If
End Sub
Function Indexing() As Collection
'インデックスをコレクションで返す
    Dim File As String
    File = Dir(Path & Application.PathSeparator & "*.xls*")
    
    Dim Books As Object
    Set Books = CreateObject("Scripting.Dictionary") 'グローバル変数のBooksコレクションインスタンスを生成
    
    Do While File <> ""
        If Not left$(File, 1) = "~" Then 'バックアップファイルや隠しファイルみたいなものは避ける
            Dim wkb As Workbook
            Dim wkbpath
            wkbpath = Path & Application.PathSeparator & File
            
            'ハッシュを比較するのは一度スキャンしたことがあるブックのみ
            'ブックのハッシュ値は独立したテーブルに保存する
            Dim sha1
            sha1 = hash.calcSha1(wkbpath)
            Debug.Print sha1
            Set wkb = EvalBookIsOpen(wkb, wkbpath) 'だれかが開いていたらすぐ閉じる
            ActiveWindow.Visible = False
            If Not wkb Is Nothing Then
                Debug.Print wkb.name
                Scan wkb, wkbpath, Books
                CloseWkb wkb
            End If
        End If
        File = Dir()
    Loop
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            Indexing (f.Path)
        Next
    End With
        
    updateDB Books

    Set wkb = Nothing
    Set f = Nothing
End Function
Private Sub Scan(wkb, wkbpath, Books)
    Dim IDs
    IDs = GetSheetNames(wkb)
    On Error Resume Next '同じキーが出てくるけど、それは関係ないシートだからスキップしてOK
    Dim v
    For Each v In IDs
        Books(v) = wkbpath
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
        If RegularExpressions.RegexTest(st.name, "^N\d{2}[A-Z]-+\d+") Then
            names(c) = st.name
            c = c + 1
        End If
    Next
    GetSheetNames = names
End Function

Private Sub updateDB(coll As Object)
        Const TABLE_NAME As String = "quotation_index"
        Const DB_FILE_NAME As String = "winckler.sqlite3"
        Dim sql As String
        sql = "INSERT INTO " & TABLE_NAME & " (quotation_id, quotation_fullpath) VALUES (" & _
                        "?, ?)"
        sqlite_no_ADODB.Insert sql, coll, DB_FILE_NAME
End Sub
Private Function EvalBookIsOpen(wkb, Path) As Workbook
'使用中のファイルを開かない
    Application.DisplayAlerts = False
    Set wkb = Workbooks.Open(fileName:=Path, Notify:=False)
    If wkb.ReadOnly Then
        wkb.Close
    End If
    Set EvalBookIsOpen = wkb
    Application.DisplayAlerts = True
End Function
