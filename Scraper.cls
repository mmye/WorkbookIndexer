VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scraper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
' Class to scrape whole information from a workbook.
' Potential purpose of this is to make index, which search program runs to find user input and returns corresponding quotation id.

Dim st As Worksheet
Dim var() As Variant
Dim vName() As Variant
Dim c As Long

Sub Init(Sheet As Worksheet)
    Set st = Sheet
End Sub
Function Scrape() As Variant
    'ゴミが入らないようになるべくシート範囲を限定したい
    Dim LastRow
    LastRow = st.UsedRange.item(st.UsedRange.Count).row
    Scrape = Intersect(st.UsedRange, st.rows("29:" & LastRow & ""))
End Function

