Attribute VB_Name = "SmartSheetCaller"
Option Explicit

Sub CallClass()

Dim st As New SmartSheet
st.Init ActiveSheet
st.Pull

End Sub
