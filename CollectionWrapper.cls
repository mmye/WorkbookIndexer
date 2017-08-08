VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Books As Collection

Private Sub Class_Initialize()
    Set Books = New Collection
End Sub

Public Function Add(ByVal node As String) As String
    Call Books.Add(node)
    Add = node
End Function

Public Sub Remove(ByVal index As Integer)
    Call Books.Remove(index + 1)
End Sub

Public Property Get Count() As Long
    Count = Books.Count
End Property

Public Property Get Item(ByVal index As Long) As String
    Item = Books(index + 1)
End Property

Private Sub Class_Terminate()
    Debug.Print ("CollectionWraper Class destructer called.")
End Sub
