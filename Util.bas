Attribute VB_Name = "Util"
Option Explicit

'***********************************************************
' �@�\   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
'***********************************************************
Public Function IsArrayEx(varArray As Variant) As Long
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
