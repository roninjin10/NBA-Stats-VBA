Attribute VB_Name = "Scraping4"
Option Explicit

Sub BuildBoxScores()
    Dim Plrs As Dictionary
    Dim plr As player: Set plr = New player
    Set Plrs = New Dictionary
    Dim i As Long: i = 0
    Dim plrname As String
    Dim rng As Range
    For Each rng In Range(wSiteCSVs.Range("B2"), wSiteCSVs.Range("B2").End(xlDown))
        plrname = rng.Value
        Call plr.CreatePlayer(plrname)
        Plrs.Add key:=i, item:=plr
        i = i + 1
    Next rng
End Sub
