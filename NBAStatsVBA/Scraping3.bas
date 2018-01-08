Attribute VB_Name = "Scraping3"
Option Explicit

Dim IE As New SHDocVw.InternetExplorer

Private Function GetHTML(url As String, pagenum As Long) As HTMLDocument
    
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
    Dim nextbuttons As IHTMLElementCollection
    Dim nextbutton As IHTMLElement
    Set nextbuttons = GetHTML.getElementsByClassName("stats-table-pagination__next")
    Dim i As Long: i = pagenum
    Do While i <> 1
        nextbuttons(Index:=0).Click
        Set GetHTML = IE.Document
        Set nextbuttons = GetHTML.getElementsByClassName("stats-table-pagination__next")
        i = i - 1
    Loop
'
'    IE.Quit
'    Set IE = Nothing
    
End Function

Sub GetHomeAway()
    Dim objNBA As NBAStats
    Set objNBA = New NBAStats
    'Worksheets("playerHOme").Activate
    Dim hme As Worksheet
'    Set hme = Worksheets.Add
'    hme.Name = "PlayerHome"
    Set hme = Worksheets("PlayerHome")
    Call objNBA.GetPlayerstats(WorksheetToPrintTo:=hme, vLocation:=Home)
    
'    Dim awy As Worksheet
'    Set awy = Worksheets.Add
'    awy.Name = "PlayerAway"
'    Set awy = Worksheets("PlayerAway")
'    Call objNBA.GetPlayerstats(WorksheetToPrintTo:=awy, vLocation:=Away)
    
    Dim hme17 As Worksheet
'    Set hme17 = Worksheets.Add
'    hme17.Name = "PlayerHome17"
    Set hme17 = Worksheets("PlayerHome17")
    Call objNBA.GetPlayerstats(WorksheetToPrintTo:=hme17, vLocation:=Home, vSeason:=2017)
    
'    Dim awy17 As Worksheet
'    Set awy17 = Worksheets.Add
'    awy17.Name = "PlayerAway17"
'    Set hme17 = Worksheets("PlayerAway17")
'    Call objNBA.GetPlayerstats(WorksheetToPrintTo:=awy17, vLocation:=Away)
End Sub
Sub ScrapePlayers()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim HTMLDoc As HTMLDocument
    
    Dim url As String: url = "http://stats.nba.com/players/transition/"
    Dim Transition As Worksheet
    Set Transition = Worksheets("plTransition")
    Transition.Activate
    Transition.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, Transition)
    
    Dim Isolation As Worksheet
    Set Isolation = Worksheets("plIsos")
    url = "http://stats.nba.com/players/isolation/"
    Isolation.Activate
    Isolation.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, Isolation)
    
    Dim PNRHandler As Worksheet
    Set PNRHandler = Worksheets("plPNRBall")
    url = "http://stats.nba.com/players/ball-handler/"
    PNRHandler.Activate
    PNRHandler.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRHandler)
    
    Dim PNRRoller As Worksheet
    Set PNRRoller = Worksheets("plPNRRoll")
    PNRRoller.Activate
    url = "http://stats.nba.com/players/ball-handler/"
    PNRRoller.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, PNRRoller)
    
    Dim Postups As Worksheet
    Set Postups = Worksheets("plPostUps")
    Postups.Activate
    url = "http://stats.nba.com/players/playtype-post-up/"
    Postups.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, Postups)
    
    Dim Spotups As Worksheet
    Set Spotups = Worksheets("plSpotUps")
    Spotups.Activate
    url = "http://stats.nba.com/players/spot-up/"
    Spotups.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, Spotups)
    
    Dim HandOffs As Worksheet
    Set HandOffs = Worksheets("plHandOffs")
    HandOffs.Activate
    url = "http://stats.nba.com/players/hand-off/"
    HandOffs.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, HandOffs)
    
    Dim Cuts As Worksheet
    Set Cuts = Worksheets("plCuts")
    Cuts.Activate
    url = "http://stats.nba.com/players/cut/"
    Cuts.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, Cuts)
    
    Dim OffScreens As Worksheet
    Set OffScreens = Worksheets("plOffScreens")
    OffScreens.Activate
    url = "http://stats.nba.com/players/off-screen/"
    OffScreens.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, OffScreens)
    
    Dim PutBacks As Worksheet
    Set PutBacks = Worksheets("plPutBacks")
    PutBacks.Activate
    url = "http://stats.nba.com/players/putbacks/"
    PutBacks.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, PutBacks)
    
    Dim Misc As Worksheet
    Set Misc = Worksheets("plMisc")
    Misc.Activate
    url = "http://stats.nba.com/players/playtype-misc/"
    Misc.Range("A:V").ClearContents
    
    Set HTMLDoc = GetHTML(url, 1)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
    Set HTMLDoc = GetHTML(url, 2)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
    Set HTMLDoc = GetHTML(url, 3)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
    Set HTMLDoc = GetHTML(url, 4)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
    Set HTMLDoc = GetHTML(url, 5)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
    Set HTMLDoc = GetHTML(url, 6)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
    Set HTMLDoc = GetHTML(url, 7)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
    Set HTMLDoc = GetHTML(url, 8)
    Call ProcessHTMLPageNBA(HTMLDoc, Misc)
'
'    IE.Quit
'    Set IE = Nothing
    
        
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Public Sub openreadonly()
    Dim wb As Workbook
    Dim DailyRotoTable As String
    
    DailyRotoTable = "C:\Users\willc\Desktop\New folder (2)\Custom_Scraper\out\DAILYROTO_TABLE"
    Set wb = Workbooks.Open(Filename:=DailyRotoTable, ReadOnly:=True, UpdateLinks:=False)
    Debug.Print ActiveSheet.Range("A2")
    wb.Close
    Set wb = Nothing
End Sub


Private Sub ProcessHTMLPageNBA(HTMLPage As MSHTML.HTMLDocument, ws As Worksheet)
        
        Dim HTMLTable As MSHTML.IHTMLElement
        Dim HTMLTables As MSHTML.IHTMLElementCollection
        Dim HTMLRow As MSHTML.IHTMLElement
        Dim HTMLCell As MSHTML.IHTMLElement
        Dim rownum As Long, colnum As Long
        Dim TableStartRange As Range
        
        'ws.UsedRange.ClearContents
        ws.Activate
        
        Set HTMLTables = HTMLPage.getElementsByTagName("table")
        
        
        For Each HTMLTable In HTMLTables
            'Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
            'On Error Resume Next
            'If TableName = HTMLTable.ID Then
                Set TableStartRange = Range("A9999").End(xlUp).Offset(1, 0)
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
            Exit For
        Next HTMLTable
        
        If Application.WorksheetFunction.CountIf(Range(ws.Range("A1"), TableStartRange.Offset(-1, 0)), TableStartRange.Offset(3, 0).Value) <> 0 Then
            Range(TableStartRange, ws.Range("U9999")).ClearContents
        End If

        
    End Sub
