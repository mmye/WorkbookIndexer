Attribute VB_Name = "WorkbookScraper"
Option Explicit

Sub Scrape()
    Dim s As Scraper
    Set s = New Scraper
    Dim st As Worksheet
    For Each st In ActiveWorkbook.Sheets
        s.Init st
        s.Scrape
    Next
    Stop
End Sub
