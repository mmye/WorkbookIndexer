Attribute VB_Name = "DynamicGenerationControls"
Option Explicit

Sub Dynamic()

Dim myForm As Control


Dim cCntrl As Control
Dim tp As Double

tp = 25
For i = 1 To 5
    With Me.Controls
        Set cCntrl = .Add("Forms.textbox.1")
        With cCntrl
            .Width = 150
            .Height = 25
            .Top = tp
            .Left = 10
            .ZOrder (0)
        End With
        tp = tp + 25
    End With
Next i

End Sub

