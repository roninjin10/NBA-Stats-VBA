Attribute VB_Name = "Scraping2"
Option Explicit

Private Const ChiRotoWorld = "http://www.rotoworld.com/teams/nba/chi/chicago-bulls"

Public Sub GetTeamNews()
    Dim team As String, RotoworldLink As String
    team = ActiveCell.Value
    Dim Teams As Range: Set Teams = Range(wTeamInfo.Range("A2"), wTeamInfo.Range("A2").End(xlDown))
    Dim RotoLinks As Range: Set RotoLinks = Range(wTeamInfo.Range("D2"), wTeamInfo.Range("D2").End(xlDown))
    RotoworldLink = RotoLinks.Cells(BinarySearchString(team, Teams)).Value
    
    Call ParseTeamNews(RotoworldLink)
    
End Sub

Private Function GetHTML(url As String) As HTMLDocument
    Dim IE As New SHDocVw.InternetExplorer
    Dim iretry As Long: iretry = 5
    If IE Is Nothing Then Set IE = New InternetExplorer
    
    IE.navigate url
    IE.Visible = False
    
    While IE.Busy: DoEvents: Wend
    While IE.readyState <> READYSTATE_COMPLETE: DoEvents: Wend
    
    Application.Wait DateAdd("s", 2, Now)
    
    Dim ieDoc As HTMLDocument: Set ieDoc = IE.Document
    Application.Wait DateAdd("s", 1, Now)
    DoEvents

    Dim ieDiv As HTMLDivElement
    Set ieDiv = Nothing
    While ieDiv Is Nothing And iretry < 5
        iretry = iretry + 1
        Set ieDiv = ieDoc.getElementsByClassName("pb")(0): DoEvents
    Wend
    
    Set GetHTML = New MSHTML.HTMLDocument
    Set GetHTML = IE.Document
    'Debug.Print IE.Document
    
'    IE.Quit
'    Set IE = Nothing
    
End Function

Private Sub ParseTeamNews(url As String)
    Dim i As Long: i = 0
    Dim HTMLDoc As New MSHTML.HTMLDocument
    Set HTMLDoc = GetHTML(ChiRotoWorld)
    Dim ReportDic As New Dictionary
    Dim ImpactDic As New Dictionary
    
    Dim report As MSHTML.IHTMLElement
    Dim reports As MSHTML.IHTMLElementCollection
    Set reports = HTMLDoc.getElementsByClassName("report")
    For Each report In reports
        i = i + 1
        ReportDic.Add key:=i, item:=report.innerText
        'Debug.Print report.innerText
    Next report
    Dim TotalReports As Long: TotalReports = i
    
    Dim Impact As MSHTML.IHTMLElement
    Dim Impacts As MSHTML.IHTMLElementCollection
    Set Impacts = HTMLDoc.getElementsByClassName("impact")
    Dim j As Long: j = 0
    For Each Impact In Impacts
        j = j + 1
        ImpactDic.Add key:=j, item:=Impact.innerText
        'Debug.Print Impact.innerText
    Next Impact
    Dim TotalImpacts As Long: TotalImpacts = j
    
    
    
    Dim ws As Worksheet
    Set ws = Worksheets.Add
    ws.Range("A1") = "ReportNumber"
    ws.Range("B1") = "Report"
    ws.Range("C1") = "Impact"
    
    
    Dim key As Variant
    Dim TeamReport As String
    For key = 1 To TotalReports
        ws.Cells(key + 1, 1).Value = key
        ws.Cells(key + 1, 2).Value = ReportDic(key)
        ws.Cells(key + 1, 3).Value = ImpactDic(key + 1) 'first impact report is instructions "to search for a player use one of two formats:"
    Next key
End Sub

Public Sub UpdateTeamPlaytypes()
    Dim URLs As New Dictionary
    Dim ws As New Dictionary
    Dim HTMLDoc As New MSHTML.HTMLDocument
    

    URLs.Add key:=1, item:="http://stats.nba.com/teams/transition/"
    ws.Add key:=1, item:=wOffTrans
    URLs.Add key:=2, item:="http://stats.nba.com/teams/transition/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=2, item:=wDefTransition
    URLs.Add key:=3, item:="http://stats.nba.com/teams/isolation/"
    ws.Add key:=3, item:=wOffIsos
    URLs.Add key:=4, item:="http://stats.nba.com/teams/isolation/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=4, item:=wDefIsos
    URLs.Add key:=5, item:="http://stats.nba.com/teams/ball-handler/"
    ws.Add key:=5, item:=wOfPNRBallHandler
    URLs.Add key:=6, item:="http://stats.nba.com/teams/ball-handler/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=6, item:=wDefPNRBallHandler
    URLs.Add key:=7, item:="http://stats.nba.com/teams/roll-man/"
    ws.Add key:=7, item:=wOffPNRRollMan
    URLs.Add key:=8, item:="http://stats.nba.com/teams/roll-man/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=8, item:=wDefPNRRollMan
    URLs.Add key:=9, item:="http://stats.nba.com/teams/playtype-post-up/"
    ws.Add key:=9, item:=wOffPostups
    URLs.Add key:=10, item:="http://stats.nba.com/teams/playtype-post-up/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=10, item:=wDefPostups
    URLs.Add key:=11, item:="http://stats.nba.com/teams/spot-up/"
    ws.Add key:=11, item:=wOffSpotups
    URLs.Add key:=12, item:="http://stats.nba.com/teams/spot-up/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=12, item:=wDefSpotups
    URLs.Add key:=13, item:="http://stats.nba.com/teams/hand-off/"
    ws.Add key:=13, item:=wOffHandoffs
    URLs.Add key:=14, item:="http://stats.nba.com/teams/hand-off/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=14, item:=wDefHandoffs
    URLs.Add key:=15, item:="http://stats.nba.com/teams/cut/"
    ws.Add key:=15, item:=wOffCuts
    URLs.Add key:=16, item:="http://stats.nba.com/teams/cut/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=16, item:=wDefCuts
    URLs.Add key:=17, item:="http://stats.nba.com/teams/off-screen/"
    ws.Add key:=17, item:=wOffOffScreens
    URLs.Add key:=18, item:="http://stats.nba.com/teams/off-screen/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=18, item:=wDefOffScreens
    URLs.Add key:=19, item:="http://stats.nba.com/teams/putbacks/"
    ws.Add key:=19, item:=wOffPutbacks
    URLs.Add key:=20, item:="http://stats.nba.com/teams/putbacks/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=20, item:=wDefPutbacks
    URLs.Add key:=21, item:="http://stats.nba.com/teams/playtype-misc/"
    ws.Add key:=21, item:=wOffMisc
    URLs.Add key:=22, item:="http://stats.nba.com/teams/playtype-misc/?Season=2017-18&SeasonType=Regular%20Season&PerMode=Totals&OD=defensive"
    ws.Add key:=22, item:=wDefMisc
    
    Dim ws2 As Worksheet
    Dim key As Variant
    For Each key In URLs.Keys
        Set HTMLDoc = GetHTML(URLs(key))
        Set ws2 = ws(key)
        Call ProcessHTMLPageNBA(HTMLDoc, ws2)
    Next key
End Sub

Private Sub ProcessHTMLPageNBA(HTMLPage As MSHTML.HTMLDocument, ws As Worksheet)
        
        Dim HTMLTable As MSHTML.IHTMLElement
        Dim HTMLTables As MSHTML.IHTMLElementCollection
        Dim HTMLRow As MSHTML.IHTMLElement
        Dim HTMLCell As MSHTML.IHTMLElement
        Dim rownum As Long, colnum As Long
        Dim TableStartRange As Range
        
        ws.Range("A:T").ClearContents
        
        Set HTMLTables = HTMLPage.getElementsByTagName("table")
        
        For Each HTMLTable In HTMLTables
            'Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
            'On Error Resume Next
            'If TableName = HTMLTable.ID Then
                If Application.WorksheetFunction.CountA(Range("A1")) = 1 Then
                    Set TableStartRange = ws.Range("A99999").End(xlUp).Offset(1, 0)
                    If ws.Range("A1").End(xlDown).Row < ws.Range("B1").End(xlDown).Row Then
                        Set TableStartRange = ws.Range("B1").End(xlDown).Offset(1, -1)
                    End If
                Else
                    Set TableStartRange = Range("A1")
                End If
                'TableStartRange.Value = HTMLTable.ID
                TableStartRange.Offset(0, 1) = Now
                rownum = 2
                For Each HTMLRow In HTMLTable.getElementsByTagName("tr")
                    colnum = 1
                    For Each HTMLCell In HTMLRow.Children
                        TableStartRange.Offset(rownum - 1, colnum - 1).Value = HTMLCell.innerText
                        colnum = colnum + 1
                    Next HTMLCell
                    rownum = rownum + 1
                Next HTMLRow
            'End If
        Next HTMLTable


    End Sub
