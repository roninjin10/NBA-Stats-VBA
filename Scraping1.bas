Attribute VB_Name = "Scraping1"
Option Explicit

Dim IE As New SHDocVw.InternetExplorer

Sub UpdateAllStats()
    Dim starttime As Variant
    Dim endtime As Variant
    
    MsgBox "updatingStats"
    
    starttime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call GetOpponentStats
    Call UpdateAdvancedTeamLast3
    Call UpdatePaceAdjustedTeam2018
    Call UpdateAdvancedTeam2018
    Call updateBasicTeam2018
    Call UpdatePaceAdj2018
    Call UPdateAdvanced2018
    Call UPdateBasic2018
    Call UpdateLast5Advanced
    Call UpdateLast5
    Call UpdateLastGameAdvanced
    Call UpdateLastGame
    Call BuildLastxGames
    Call UpdateTeamPlaytypes
    Call SetupMOdel
    Call AliasPlayers
    
    Dim rng As Range
    For Each rng In Range(wAdvanced2018.Range("B2"), wAdvanced2018.Range("B2").End(xlDown))
        If Alias(rng.Value) = "" Then
            MsgBox rng.Value & " Has No Alias"
        End If
    Next rng
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    endtime = Timer
    
    MsgBox endtime - starttime
    
End Sub


Sub AliasPlayers()
    Dim player As Range
    Dim Players As Range
    Dim ws As Variant
    Dim wss() As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ReDim wss(1 To 10)
    Set wss(1) = Worksheets("PaceAdj2018")
    Set wss(2) = Worksheets("Advanced2018")
    Set wss(3) = Worksheets("Basic2018")
    Set wss(4) = Worksheets("last5adv")
    Set wss(5) = Worksheets("last5")
    Set wss(6) = Worksheets("lastgameadv")
    Set wss(7) = Worksheets("lastgame")
    Set wss(8) = Worksheets("PaceAdj2017")
    Set wss(9) = Worksheets("Advanced2017")
    Set wss(10) = Worksheets("Basic2017")
    
    Dim AliasValue As String
    For Each ws In wss
        Set Players = Range(ws.Range("B2"), ws.Range("B2").End(xlDown))
        For Each player In Players
            AliasValue = Alias(player.Value)
            If AliasValue <> "" Then
                player.Value = Alias(player.Value)
            Else
                'MsgBox Player.Value
            End If
        Next player
    Next ws
    
    Set ws = wFantasyLabs
    Set Players = Range(ws.Range("A2"), ws.Range("A2").End(xlDown))
    For Each player In Players
        AliasValue = Alias(player.Value)
        If AliasValue <> "" Then
            player.Value = Alias(player.Value)
        Else
'            MsgBox player.Value
        End If
    Next player
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
End Sub

Sub BuildLastxGames()
    Dim objNBA As NBAStats
    Set objNBA = New NBAStats
    
    Dim Lastx As Long
    For Lastx = 1 To 10
        Dim ws1 As Worksheet
        Set ws1 = Worksheets("Last" & Lastx)
        ws1.UsedRange.ClearContents
        'ws1.Name = "Last" & Lastx
        Call objNBA.GetPlayerstats(WorksheetToPrintTo:=ws1, vMeasureType:=basic, vPaceAdjust:=paceno, vLastNGames:=str(Lastx))
        Dim ws2 As Worksheet
        Set ws2 = Worksheets("Last" & Lastx & "Adv")
        ws2.UsedRange.ClearContents
        'ws2.Name = "Last" & Lastx & "Adv"
        Call objNBA.GetPlayerstats(WorksheetToPrintTo:=ws2, vMeasureType:=advanced, vPaceAdjust:=paceno, vLastNGames:=str(Lastx))
        Dim ws3 As Worksheet
        Set ws3 = Worksheets("Last" & Lastx & "pa")
        ws3.UsedRange.ClearContents
        'ws3.Name = "Last" & Lastx & "pa"
        Call objNBA.GetPlayerstats(WorksheetToPrintTo:=ws3, vMeasureType:=basic, vPaceAdjust:=paceyes, vLastNGames:=str(Lastx))
    Next Lastx
End Sub

Sub GetOpponentStats()

    Dim ws As Worksheet
    
    Set ws = wBBallRefOpponent
    Debug.Print ws.Name
    ws.Activate
    ws.Range("A1").Activate

    ws.Range("B2").CurrentRegion.ClearContents
    ws.Range("A1").Select
    
    Call ProcessHTMLPageBBallRef(GetHTMLBBallRef("https://www.basketball-reference.com/leagues/NBA_2018.html"), "opponent-stats-per_game")

    'Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
    
End Sub
Function GetHTMLBBallRef(url As String) As MSHTML.HTMLDocument
        
        Dim HTMLDoc As New MSHTML.HTMLDocument
        Dim HTMLTables As MSHTML.IHTMLElementCollection, HTMLTable As MSHTML.IHTMLElement
        Dim TableNumber As Long, LastTableNumber As Long
        
        'URL = "https://www.basketball-reference.com/leagues/NBA_2018.html"
        
        IE.Visible = False
        IE.navigate url
        Do While IE.readyState <> READYSTATE_COMPLETE
        Loop
        
        Set HTMLDoc = IE.Document
        Set HTMLTables = HTMLDoc.getElementsByTagName("Table")
        TableNumber = 0
        LastTableNumber = -1
        
        Do While TableNumber <> LastTableNumber
            Application.Wait Now + TimeValue("0:00:01")
            LastTableNumber = TableNumber
            Set HTMLDoc = IE.Document
            Set HTMLTables = HTMLDoc.getElementsByTagName("Table")
            TableNumber = HTMLTables.Length
            Debug.Print " "
        Loop
        
        Set HTMLDoc = IE.Document
        
        Set GetHTMLBBallRef = HTMLDoc
        
'        IE.Quit
'        Set IE = Nothing
        
    End Function
    
    
    Sub ProcessHTMLPageBBallRef(HTMLPage As MSHTML.HTMLDocument, TableName As String)
        
        Dim HTMLTable As MSHTML.IHTMLElement
        Dim HTMLTables As MSHTML.IHTMLElementCollection
        Dim HTMLRow As MSHTML.IHTMLElement
        Dim HTMLCell As MSHTML.IHTMLElement
        Dim rownum As Long, colnum As Long
        Dim TableStartRange As Range
        Dim ws As Worksheet

        
        Set HTMLTables = HTMLPage.getElementsByTagName("table")
        Set ws = wBBallRefOpponent
        
        For Each HTMLTable In HTMLTables
            'Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
            On Error Resume Next
            If TableName = HTMLTable.ID Then
                If Application.WorksheetFunction.CountA(Range("A1")) = 1 Then
                    Set TableStartRange = ws.Range("A1").End(xlDown).Offset(1, 0)
                    If ws.Range("A1").End(xlDown).Row < ws.Range("B1").End(xlDown).Row Then
                        Set TableStartRange = ws.Range("B1").End(xlDown).Offset(1, -1)
                    End If
                Else
                    Set TableStartRange = Range("A1")
                End If
                TableStartRange.Value = HTMLTable.ID
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
            End If
        Next HTMLTable


    End Sub
    
    Sub UpdateAdvancedTeamLast3()
        Dim ws As Worksheet
        Dim objNBA As NBAStats
        
        Set objNBA = New NBAStats
        Set ws = wAdvancedTeamLast3
        ws.Activate
        ws.Range("B2").CurrentRegion.ClearContents
        
        Call objNBA.GetTeamStats(WorksheetToPrintTo:=ws, vMeasureType:=advanced, vLastNGames:="3")
        
        
    End Sub
    
'    Sub UpdateAdvancedTeamLast10()
'        Dim ws As Worksheet
'        Dim objNBA As NBAStats
'
'        Set objNBA = New NBAStats
'        worsheets.Add
'        activeworksheet.Name = "AdvancedTeamLast5"
'
'        Set ws = wAdvancedTeamLast10
'        ws.Range("B2").CurrentRegion.ClearContents
'
'        Call objNBA.GetTeamStats(WorksheetToPrintTo:=ws, vMeasureType:=Advanced, vLastNGames:="10")
'
'
'    End Sub
    Sub UpdatePaceAdjustedTeam2018()
        Dim ws As Worksheet
        Dim objNBA As NBAStats
        
        Set objNBA = New NBAStats
        Set ws = wPaceAdjustedTeam2018
        ws.Range("B2").CurrentRegion.ClearContents
        
        Call objNBA.GetTeamStats(WorksheetToPrintTo:=ws, vMeasureType:=basic, vLastNGames:="0", vPaceAdjust:=paceyes)
        
        
    End Sub
    
    Sub UpdateAdvancedTeam2018()
        Dim ws As Worksheet
        Dim objNBA As NBAStats
        
        Set objNBA = New NBAStats
        Set ws = wAdvancedTeam2018
        ws.Range("B2").CurrentRegion.ClearContents
        
        Call objNBA.GetTeamStats(WorksheetToPrintTo:=ws, vMeasureType:=advanced, vLastNGames:="0")  ', vPaceAdjust:=paceyes)
        
        
    End Sub
    
    
    Sub updateBasicTeam2018()
        Dim ws As Worksheet
        Dim objNBA As NBAStats
        
        Set objNBA = New NBAStats
        Set ws = wBasicTeam2018
        ws.Range("B2").CurrentRegion.ClearContents
        
        Call objNBA.GetTeamStats(WorksheetToPrintTo:=ws, vMeasureType:=basic, vLastNGames:="0")  ', vPaceAdjust:=paceyes)
        
        
    End Sub
    
Sub UpdatePaceAdj2018()
    Dim ws As Worksheet
    
    Set ws = wPaceAdj2018
    ws.Range("B2").CurrentRegion.ClearContents
    
    Call GetPlayerstats(ws:=ws, PaceAdjust:=True, YYYY:=2018)
    
End Sub

Sub UpdatePaceAdj2017()
    Dim ws As Worksheet
    
    Set ws = wPaceAdj2017
    ws.Range("B2").CurrentRegion.ClearContents
    
    Call GetPlayerstats(ws:=ws, PaceAdjust:=True, YYYY:=2017)
    
End Sub


Sub UPdateAdvanced2018()
    Dim ws As Worksheet
    
    Set ws = wAdvanced2018
    ws.Range("B2").CurrentRegion.ClearContents
    
    Call GetPlayerstats(ws:=ws, advanced:=True, PaceAdjust:=False, YYYY:=2018)
    
    
End Sub


Sub UPdateBasic2018()
    Dim ws As Worksheet
    
    Set ws = wBasic2018
    ws.Range("B2").CurrentRegion.ClearContents
    
    Call GetPlayerstats(ws:=ws, advanced:=False, PaceAdjust:=False, YYYY:=2018)
    
    
End Sub

Sub UpdateLast5Advanced()
'    Dim ws As Worksheet
'
'    Set ws = wLast5Advanced
'    ws.Range("B2").CurrentRegion.ClearContents
'
'    Call GetPlayerstats(ws:=ws, advanced:=True, LastNGames:=5)
    
End Sub

Sub UpdateLast5()
    Dim ws As Worksheet
    
    Set ws = wLast5
    ws.Range("B2").CurrentRegion.ClearContents
    
    Call GetPlayerstats(ws:=ws, advanced:=False, LastNGames:=5)
    
End Sub

Sub UpdateLastGameAdvanced()
    Dim ws As Worksheet
    
    Set ws = wLastGameAdv
    ws.Range("B2").CurrentRegion.ClearContents
    
    Call GetPlayerstats(ws:=ws, advanced:=True, LastNGames:=1)
    
End Sub

Sub UpdateLastGame()
    Dim ws As Worksheet
    
    Set ws = wLastGame
    ws.Range("B2").CurrentRegion.ClearContents
    
    Call GetPlayerstats(ws:=ws, advanced:=False, LastNGames:=1)
    
End Sub




