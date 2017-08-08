VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim f As UserForm

Sub Init(myForm As UserForm)
Stop
    Set f = myForm
    f.Caption = "This is a dynamic instance"
End Sub
