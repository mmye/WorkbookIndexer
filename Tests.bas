Attribute VB_Name = "Tests"
Option Explicit


Sub TestScrape()
    Dim v As Variant
    v = Scrape(ActiveSheet)
    Stop
End Sub
Function Scrape(st As Worksheet) As Variant
    'アクティブシートの中身をそうざらいにする
    'ただし空セルも一要素としてカウントされる
    Dim c As Scraper
    Set c = New Scraper
    
    Dim v As Variant
    c.Init st
    v = c.Scrape(v)
    Scrape = v
End Function

Sub TestArrayUtilCompress()
    Dim v As Variant
    v = Scrape(ActiveSheet)
    
    Dim c As ArrayUtil
    Set c = New ArrayUtil
    Dim ret
    ret = c.Compress(v)
Stop
End Sub

Sub TestDirLooper()
    Dim p
    p = "M:\◆事務\《1》見積・注文\1. 見積書\見積書2017"
    Dim c As DirLooper
    Set c = New DirLooper
    c.Init p
    Stop
    'BookDataにはItem=フルパス、Key=見積書番号のコレクションが返ってくる
    'Keyを参照するにはFor eachループを使う。
    Dim BookData As Collection
    Set BookData = New Collection
    Set BookData = c.Indexing
    Stop
End Sub

