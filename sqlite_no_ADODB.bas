Attribute VB_Name = "sqlite_no_ADODB"
Option Explicit

Public Const DLL_DIR As String = "C:\Users\m.maeyama\Dropbox\sqlite_for_Excel"
Public Const DB_FILE_DIR As String = "C:\Users\m.maeyama\Dropbox\sqlite_for_Excel\db"

Public Function SearchAll(sql As String, DB_FILE_NAME As String) As Variant
    Dim InitReturn As Long  'SQLiteDLL
    Dim dbFile As String    'DB�t�@�C��
    Dim RetVal As Long      'DBData

    #If Win64 Then
        Debug.Print "Excel is 64bit"
        'SQLiteDLL�ǂݍ���
        InitReturn = SQLite3Initialize(DLL_DIR + "\x64")
        '�f�[�^�x�[�X�n���h����`
        Dim myDbHandle As LongPtr
        '�X�e�[�g�����g�n���h����`
        Dim myStmtHandle As LongPtr
    #Else
        Debug.Print "Excel is 32bit"
        'SQLiteDLL�ǂݍ���
        InitReturn = SQLite3Initialize(DLL_DIR + "\")
        '�f�[�^�x�[�X�n���h����`
        Dim myDbHandle As Long
        '�X�e�[�g�����g�n���h����`
        Dim myStmtHandle As Long
    #End If

    'SQLiteDLL�ǂݍ��ݐ���
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & ERR.LastDllError
        Exit Function
    End If

    '�t�@�C�����擾
    dbFile = DB_FILE_DIR + "\" + DB_FILE_NAME

    'DB�t�@�C��Open
    RetVal = SQLite3Open(dbFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    Debug.Print myDbHandle

     'SQL statement�쐬
    RetVal = SQLite3PrepareV2(myDbHandle, sql, myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal

    'SQL���s
    RetVal = SQLite3Step(myStmtHandle)
    Debug.Print "SQLite3Step returned " & RetVal

    '���R�[�h�擾
    Dim rows As Variant
    rows = GetRows(myStmtHandle, rows)
    
    'statement �폜
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    'DB�t�@�C��Close
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal
    
    SearchAll = rows
End Function
Private Function GetRows(dbh, arr) As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim colType As Long
    Dim rows() As Variant
    Dim ret As Long
    Dim i As Long

    rowCount = 0
    'Recordset�I�u�W�F�N�g��GetRows���\�b�h�̎擾�Ɠ���
    'rows = ***.GetRows()
    Do While ret <> SQLITE_DONE
        If rowCount = 0 Then
            colCount = SQLite3ColumnCount(dbh)
            ReDim rows(colCount - 1, rowCount)
        Else
            ReDim Preserve rows(colCount - 1, rowCount)
        End If
        For i = 0 To colCount - 1
            colType = SQLite3ColumnType(dbh, i)
            rows(i, rowCount) = ColumnValue(dbh, i, colType)
        Next
        ret = SQLite3Step(dbh)
        rowCount = rowCount + 1
    Loop
    GetRows = rows
End Function
'--------------------
'Sqlite3Demo.bas���
Private Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function
Private Sub Write2Sheet(rows As Variant)
    Dim st As Worksheet
    Set st = ActiveSheet
    Dim w As Long, h As Long
    w = UBound(rows, 2) + 1
    h = UBound(rows, 1) + 1
    
    'Row�I�u�W�F�N�g�̕��ƍ���
    Debug.Print w
    Debug.Print h
    
    Set st = Nothing
End Sub
Private Function RotateRow(rows) As Variant
    
    Dim r As Long, c As Long
    Dim i As Long
    Dim b As String
    
    Dim w As Long, h As Long
    w = UBound(rows, 2) + 1
    h = UBound(rows, 1) + 1
    
    Dim arr() As Variant
    ReDim arr(w, h) As Variant ' �^�e�ƃ��R���t�ɂ��Ă�

    'For r = LBound(Rows, 1) To UBound(Rows, 1)
    '    For c = LBound(Rows, 2) To UBound(Rows, 2)
    '        b = Rows(r, c)
    '
End Function

Public Function Insert(sql As String, rows As Object, DB_FILE_NAME As String) As Variant
    Dim InitReturn As Long  'SQLiteDLL
    Dim dbFile As String    'DB�t�@�C��
    Dim RetVal As Long      'DBData

    #If Win64 Then
        Debug.Print "Excel is 64bit"
        'SQLiteDLL�ǂݍ���
        InitReturn = SQLite3Initialize(DLL_DIR + "\x64")
        '�f�[�^�x�[�X�n���h����`
        Dim myDbHandle As LongPtr
        '�X�e�[�g�����g�n���h����`
        Dim myStmtHandle As LongPtr
    #Else
        Debug.Print "Excel is 32bit"
        'SQLiteDLL�ǂݍ���
        InitReturn = SQLite3Initialize(DLL_DIR + "\")
        '�f�[�^�x�[�X�n���h����`
        Dim myDbHandle As Long
        '�X�e�[�g�����g�n���h����`
        Dim myStmtHandle As Long
    #End If

    'SQLiteDLL�ǂݍ��ݐ���
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & ERR.LastDllError
        Exit Function
    End If

    '�t�@�C�����擾
    dbFile = DB_FILE_DIR + "\" + DB_FILE_NAME

    'DB�t�@�C��Open
    RetVal = SQLite3Open(dbFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    Debug.Print myDbHandle

    'SQL statement�쐬
    sql = SQLBuilder(sql, rows) '�����s����C��INSERT����SQL�N�G����g�ݗ��Ă�
    Debug.Print sql
    RetVal = SQLite3PrepareV2(myDbHandle, sql, myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal

    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If

    'statement �폜
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    'DB�t�@�C��Close
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal
End Function
Private Function SQLBuilder(sql, rows) As Variant
    Dim key As Variant
    For Each key In rows.keys
        Dim item As String
        item = "(" & """" & key & """" & "," & """" & rows(key) & """" & ")"
        Dim stack As String
        stack = stack & item & ","
    Next
    If Right$(stack, 1) = "," Then stack = left$(stack, Len(stack) - 1)
    Dim placeholderPos As Long
    placeholderPos = InStr(sql, "?")
        
    '�Ȃ������s���܂���̂ŁA�G���[����̂��߂ɏ�������
    stack = RegularExpressions.RegexReplace(stack, "\n", "")
    SQLBuilder = Mid$(sql, 1, placeholderPos - 2) & stack
End Function
