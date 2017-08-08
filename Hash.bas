Attribute VB_Name = "Hash"
Option Explicit

Sub Sha1Test()
    Dim BinaryStream As ADODB.stream
    Set BinaryStream = New ADODB.stream
    BinaryStream.Type = adTypeBinary
    BinaryStream.Open
    BinaryStream.LoadFromFile "C:\Users\mokoo\Dropbox\FormWindow.frm"
    
    Dim sha1 As Object
    Set sha1 = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    sha1.ComputeHash_2 (BinaryStream.Read)
    Dim hash As Variant
    hash = sha1.hash
    
    Dim i As Integer
    Dim b As String
    For i = LBound(hash) To UBound(hash)
        b = b & WorksheetFunction.Dec2Hex(hash(i), 2)
    Next i
    Debug.Print b
    BinaryStream.Close
    Set BinaryStream = Nothing
    Set sha1 = Nothing
End Sub

Function calcSha1(Path) As String
' Path�Ŏ󂯂��p�X�ɂ���t�@�C����Sha1�n�b�V����Ԃ�

    If Path = "" Then Exit Function
    ' �Q�Ɛݒ�FMicrosoft Active X Data Object x.x Librasy�@�K�v
    Dim BinaryStream As ADODB.stream
    Set BinaryStream = New ADODB.stream
    BinaryStream.Type = adTypeBinary
    BinaryStream.Open
    BinaryStream.LoadFromFile Path
    
    Dim sha1 As Object
    Set sha1 = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    sha1.ComputeHash_2 (BinaryStream.Read)
    Dim hash As Variant
    hash = sha1.hash
    
    Dim i As Integer
    Dim h As String
    For i = LBound(hash) To UBound(hash)
        h = h & WorksheetFunction.Dec2Hex(hash(i), 2)
    Next i
    BinaryStream.Close
    Set BinaryStream = Nothing
    Set sha1 = Nothing
    calcSha1 = h
End Function

