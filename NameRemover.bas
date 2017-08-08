Attribute VB_Name = "NameRemover"
Option Explicit

Sub Remove()
    Dim st As Worksheet
    For Each st In ActiveWorkbook.Sheets
        Dim n As name
        For Each n In st.names
            n.Delete
        Next
    Next
    
    For Each n In ActiveWorkbook.names
        n.Delete
    Next
    
End Sub
