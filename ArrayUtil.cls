VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim v As Variant
Function Compress(var As Variant) As Variant
    Dim ret
    ret = IsArrayEx(var)
    Select Case ret
        Case 1
            'OK!
        Case Else
            'Not OK!
            Exit Function
    End Select
    
    Dim Compressed
    Dim arr As Variant
    Dim i, c
    Const s As String = "|"
    On Error GoTo INDEX_LIMIT
    For i = LBound(var, 1) To UBound(var, 1)
        For c = LBound(var, 2) To UBound(var, 2)
            Dim d As Boolean
            d = IsDuplicate(var(i, c), var(i, c + 1))
            If Not d Then arr = arr & s & var(i, c)
        Next c
        If arr <> "" Then Compressed = arr & vbCrLf & Compressed
    Next i

    If Compressed <> "" Then Compress = Compressed
    Exit Function
INDEX_LIMIT:
    If ERR.Number = 9 Then Resume Next
End Function

'***********************************************************
' 機能   : 引数が配列か判定し、配列の場合は空かどうかも判定する
' 引数   : varArray  配列
' 戻り値 : 判定結果（1:配列/0:空の配列/-1:配列じゃない）
'***********************************************************
Private Function IsArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If ERR.Number = 9 Then
        IsArrayEx = 0
    End If
End Function
Private Function IsDuplicate(ValueA, ValueB) As Boolean
    'セルがマージされていると、同じ値が重複する。
    '引数で受けた配列のとなりあう二つの値の一致・不一致を確認して同じだったらTrueを返す
    If ValueA = Empty Then
        IsDuplicate = True
        Exit Function
    End If
    If ValueA = ValueB Then
        IsDuplicate = True
    Else
        IsDuplicate = False
    End If
End Function

