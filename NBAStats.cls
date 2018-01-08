VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NBAStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pStartPeriod As Period
Private pEndPeriod As Period
Private pGameNumber As Long
Private pSeasonYear As pSeason
Private pBoxScoreType As BoxScoreType




Private Const StartRange = "0" 'no idea what these parameters do
Private Const EndRange = "0"
Private Const RangeType = "0"

Private Const BaseTeamVsOpponent = "http://stats.nba.com/stats/teamdashboardbyopponent/?"
Private Const Base = "http://stats.nba.com/stats/"
Private Const BasePlayerStats = "http://stats.nba.com/stats/leaguedashplayerstats/?"

Private Const SeasonSegment = ""

Private Const VsConference = ""

Private Const VsDivision = ""


Public Enum Period '(_DefaultZero):
    AllQuarters = 0 '0'
    FirstQuarter = 1 '1'
    SecondQuarter = 2 '2'
    ThirdQuarter = 3 '3'
    FourthQuarter = 4 '4'
    Overtime = 5 'these weren't originally in the python version.  they had a definition for overtime as N as 4+n
    DoubleOT = 6
    TripleOT = 7 'overtimes not tested
    QuadOT = 8
End Enum

Public Enum BoxScoreType 'each type of box score avaialb efrom bball ref
    'Standard = 0  'it appears non version 2 box scores are no longer available
    'Advanced = 1
    Advancedv2 = 2
    'FourFactors = 3
    FourFactorsv2 = 4
    'Misc = 5
    Miscv2 = 6
    Playertrackv2 = 7
    'Scoring = 8
    Scoringv2 = 9
    Summuryv2 = 10
    Traditionalv2 = 11
    'Usage = 12
    Usagev2 = 13
End Enum

Public Enum pSeason
    Season2018 = 17
    Season2017 = 16
    Season2016 = 15
    Season2015 = 14
    Season2014 = 13
    Season2013 = 12
    Season2012 = 11
    Season2011 = 10
    Season2010 = 9
End Enum

Public Enum BoxScoreTypes
    boxscoreadvancedv2 = 1
    boxscorefourfactorsv2 = 2
    boxscoremiscv2 = 3
    boxscoreplayertrackv2 = 6
    boxscorescoringv2 = 5
    boxscoresummaryv2 = 4
    boxscoretraditionalv2 = 0
    boxscoreusagev2 = 7
End Enum


Public Enum ParGameScope
    season = 1
    Last10 = 2
    Yesterday = 3
    DefaultSeason = 0
End Enum

Public Enum ParPlayerExperience
    rookie = 1
    Sophomore = 2
    Veteran = 3
    DefaultNone = 0
End Enum

Public Enum ParPlayerPosition
    Forward = 1
    Center = 2
    Guard = 3
    DefaultNone = 0
End Enum

Public Enum ParStarterBench
    Starters = 1
    Bench = 2
    DefaultNone = 0
End Enum

Public Enum ParMeasureType
    basic = 1
    advanced = 2
    Misc = 3
    FourFactors = 4
    Scoring = 5
    opponent = 6
    Usage = 7
    DefaultBasic = 0
End Enum

Public Enum ParPerMode
    Totals = 1
    PerGame = 2
    MinutesPer = 3
    Per48 = 4
    Per36 = 5
    PerMinute = 6
    PerPosession = 7
    PerPlay = 8
    Per100Posessions = 9
    Per100Plays = 10
    DefaultPerGame = 0
End Enum

Public Enum ParPlusMinus
    plusYes = 1
    plusNo = 2
    DefaultNo = 0
End Enum

Public Enum ParPaceAdjust
    paceyes = 1
    paceno = 2
    DefaultNo = 0
End Enum

Public Enum ParRank
    rankYes = 1
    rankNo = 2
    DefaultNo = 0
End Enum

Public Enum ParSeasonType
    RegularSeason = 1
    Playoffs = 2
    DefaultNone = 0
End Enum

Public Enum ParOutcome
    Win = 1
    Loss = 2
    DefaultNone = 0
End Enum

Public Enum ParLocation
    Home = 1
    Away = 2
    DefaultNone = 0
End Enum

Public Enum ParMonth
    October = 1
    November = 2
    December = 3
    January = 4
    February = 5
    March = 6
    April = 7
    May = 8
    June = 9
    July = 10
    August = 11
    September = 12
    DefaultNone = 0
End Enum

Public Enum ParGameSegment
    FirstHalf = 1
    SecondHalf = 2
    Overtime = 3
    DefaultAll = 0
End Enum

Public Enum ParPeriod
    First = 1
    Second = 2
    Third = 3
    Fourth = 4
    DefaultAll = 0
End Enum


Public Property Get StartPeriod() As Period
    StartPeriod = (pStartPeriod)
End Property

Public Property Let StartPeriod(NewStartPeriod As Period)
    pStartPeriod = NewStartPeriod
End Property

Public Property Get EndPeriod() As Period
    EndPeriod = (pEndPeriod)
End Property

Public Property Let EndPeriod(NewEndPeriod As Period)
    pEndPeriod = NewEndPeriod
End Property

Public Property Get gamenumber() As Long
    gamenumber = pGameNumber
End Property

Public Property Let gamenumber(NewGameNumber As Long)
        pGameNumber = NewGameNumber
End Property

Public Property Get SeasonYear() As pSeason
    SeasonYear = pSeasonYear
End Property

Public Property Let SeasonYear(NewSeasonYear As pSeason)
    pSeasonYear = NewSeasonYear
End Property

Public Property Get BoxScoreCategory() As BoxScoreType
    BoxScoreCategory = pBoxScoreType
End Property

Public Property Let BoxScoreCategory(NewBoxScoreCategory As BoxScoreType)
    pBoxScoreType = NewBoxScoreCategory
End Property


Private Function GetGameScope(i As Long) As String
    If i = 1 Then
        GetGameScope = ""
    ElseIf i = 2 Then
        GetGameScope = "Last 10"
    ElseIf i = 3 Then
        GetGameScope = "Yesterday"
    ElseIf i = 4 Then
        GetGameScope = "Finals"
    ElseIf i = 0 Then
        GetGameScope = ""
    End If
End Function


Private Function GetPlayerExperience(i As Long) As String
    If i = 1 Then
        GetPlayerExperience = "Rookie"
    ElseIf i = 2 Then
        GetPlayerExperience = "Sophomore"
    ElseIf i = 3 Then
        GetPlayerExperience = "Veteran"
    ElseIf i = 0 Then
        GetPlayerExperience = ""
    End If
End Function



Private Function GetPlayerPosition(i As Long) As String
    If i = 1 Then
        GetPlayerPosition = "F"
    ElseIf i = 2 Then
        GetPlayerPosition = "C"
    ElseIf i = 3 Then
        GetPlayerPosition = "G"
    ElseIf i = 0 Then
        GetPlayerPosition = ""
    End If
End Function



Private Function GetStarterBench(i As Long) As String
    If i = 1 Then
        GetStarterBench = "Starters"
    ElseIf i = 2 Then
        GetStarterBench = "Bench"
    ElseIf i = 0 Then
        GetStarterBench = ""
    End If
End Function



Private Function GetMeasureType(i As Long) As String
    If i = 1 Then
        GetMeasureType = "Base"
    ElseIf i = 2 Then
        GetMeasureType = "Advanced"
    ElseIf i = 3 Then
        GetMeasureType = "Misc"
    ElseIf i = 4 Then
        GetMeasureType = "Four Factors"
    ElseIf i = 5 Then
        GetMeasureType = "Scoring"
    ElseIf i = 6 Then
        GetMeasureType = "Opponent"
    ElseIf i = 7 Then
        GetMeasureType = "Usage"
    ElseIf i = 0 Then
        GetMeasureType = "Base"
    End If
End Function



Private Function GetPerMode(i As Long) As String
    If i = 1 Then
        GetPerMode = "Totals"
    ElseIf i = 2 Then
        GetPerMode = "PerGame"
    ElseIf i = 3 Then
        GetPerMode = "MinutesPer"
    ElseIf i = 4 Then
        GetPerMode = "Per48"
    ElseIf i = 5 Then
        GetPerMode = "Per36"
    ElseIf i = 6 Then
        GetPerMode = "PerMinute"
    ElseIf i = 7 Then
        GetPerMode = "PerPosession"
    ElseIf i = 8 Then
        GetPerMode = "PerPlay"
    ElseIf i = 9 Then
        GetPerMode = "Per100Posessions"
    ElseIf i = 10 Then
        GetPerMode = "Per100Plays"
    ElseIf i = 0 Then
        GetPerMode = "PerGame"
    End If
End Function



Private Function GetPlusMinus(i As Long) As String
    If i = 1 Then
        GetPlusMinus = "Y"
    ElseIf i = 2 Then
        GetPlusMinus = "N"
    ElseIf i = 0 Then
        GetPlusMinus = "N"
    End If
End Function



Private Function GetPaceAdjust(i As Long) As String
    If i = 1 Then
        GetPaceAdjust = "Y"
    ElseIf i = 2 Then
        GetPaceAdjust = "N"
    ElseIf i = 0 Then
        GetPaceAdjust = "N"
    End If
End Function



Private Function GetRank(i As Long) As String
    If i = 1 Then
        GetRank = "Y"
    ElseIf i = 2 Then
        GetRank = "N"
    ElseIf i = 0 Then
        GetRank = "N"
    End If
End Function


Private Function GetSeason(season As Long) As String
    GetSeason = (season - 1) & "-" & Right(season, 2)
End Function


Private Function GetSeasonType(i As Long) As String
    If i = 1 Then
        GetSeasonType = "Regular Season"
    ElseIf i = 2 Then
        GetSeasonType = "Playoffs"
    ElseIf i = 0 Then
        GetSeasonType = "Regular Season"
    End If
End Function



Private Function GetOutcome(i As Long) As String
    If i = 1 Then
        GetOutcome = "W"
    ElseIf i = 2 Then
        GetOutcome = "L"
    ElseIf i = 0 Then
        GetOutcome = ""
    End If
End Function



Private Function GetLocation(i As Long) As String
    If i = 1 Then
        GetLocation = "Home"
    ElseIf i = 2 Then
        GetLocation = "Away"
    ElseIf i = 0 Then
        GetLocation = ""
    End If
End Function



Private Function GetMonth(i As Long) As String
    If i = 1 Then
        GetMonth = "1"
    ElseIf i = 2 Then
        GetMonth = "2"
    ElseIf i = 3 Then
        GetMonth = "3"
    ElseIf i = 4 Then
        GetMonth = "4"
    ElseIf i = 5 Then
        GetMonth = "5"
    ElseIf i = 6 Then
        GetMonth = "6"
    ElseIf i = 7 Then
        GetMonth = "7"
    ElseIf i = 8 Then
        GetMonth = "8"
    ElseIf i = 9 Then
        GetMonth = "9"
    ElseIf i = 10 Then
        GetMonth = "10"
    ElseIf i = 11 Then
        GetMonth = "11"
    ElseIf i = 12 Then
        GetMonth = "12"
    ElseIf i = 0 Then
        GetMonth = "0"
    End If
End Function


Public Function GetTeamID(team As String) As String
    GetTeamID = "0"
    If UCase(team) <= "MEM" Then
        If UCase(team) <= "DEN" Then
            If UCase(team) <= "CHA" Then
                Select Case team
                    Case "Atl"
                        GetTeamID = "1610612737"
                    Case "Bos"
                        GetTeamID = "1610612738"
                     Case "Bkn"
                        GetTeamID = "1610612751"
                    Case "Cha"
                        GetTeamID = "1610612766"
                End Select
            Else
                Select Case team
                    Case "Chi"
                        GetTeamID = "1610612741"
                    Case "Cle"
                        GetTeamID = "1610612739"
                    Case "Dal"
                        GetTeamID = "1610612742"
                    Case "Den"
                        GetTeamID = "1610612743"
                End Select
            End If
        Else
            If UCase(team) <= "IND" Then
                Select Case team
                    Case "Det"
                        GetTeamID = "1610612765"
                    Case "Gs"
                        GetTeamID = "1610612744"
                    Case "Hou"
                        GetTeamID = "1610612745"
                    Case "Ind"
                        GetTeamID = "1610612754"
                End Select
            Else
                Select Case team
                    Case "Lac"
                        GetTeamID = "1610612746"
                    Case "Lal"
                        GetTeamID = "1610612747"
                    Case "Mem"
                        GetTeamID = "1610612763"
                End Select
            End If
        End If
    Else
        If UCase(team) <= "PHI" Then
            If UCase(team) <= "NY" Then
                Select Case team
                    Case "Mia"
                        GetTeamID = "1610612748"
                    Case "Mil"
                        GetTeamID = "1610612749"
                    Case "Min"
                        GetTeamID = "1610612750"
                    Case "No"
                        GetTeamID = "1610612740"
                    Case "Ny"
                        GetTeamID = "1610612752"
                End Select
            Else
                Select Case team
                    Case "Okc"
                        GetTeamID = "1610612760"
                    Case "Orl"
                        GetTeamID = "1610612753"
                    Case "Phi"
                        GetTeamID = "1610612755"
                End Select
            End If
        Else
            If UCase(team) <= "SA" Then
                Select Case team
                    Case "Pho"
                        GetTeamID = "1610612756"
                    Case "Por"
                        GetTeamID = "1610612757"
                    Case "Sac"
                        GetTeamID = "1610612758"
                    Case "Sa"
                        GetTeamID = "1610612759"
                End Select
            Else
                Select Case team
                    Case "Tor"
                        GetTeamID = "1610612761"
                    Case "Uta"
                        GetTeamID = "1610612762"
                    Case "Was"
                        GetTeamID = "1610612764"
                End Select
            End If
        End If
    End If
End Function




Private Function GetGameSegment(i As Long) As String
    If i = 1 Then
        GetGameSegment = "First Half"
    ElseIf i = 2 Then
        GetGameSegment = "Second Half"
    ElseIf i = 3 Then
        GetGameSegment = "Overtime"
    ElseIf i = 0 Then
        GetGameSegment = ""
    End If
End Function




Private Function GetPeriod(i As Long) As String
    If i = 1 Then
        GetPeriod = "1"
    ElseIf i = 2 Then
        GetPeriod = "2"
    ElseIf i = 3 Then
        GetPeriod = "3"
    ElseIf i = 4 Then
        GetPeriod = "4"
    ElseIf i = 0 Then
        GetPeriod = "0"
    End If
End Function




Private Function GetBoxScoreType(item As BoxScoreType) As String
    Select Case item
        Case 0: GetBoxScoreType = "boxscore"
        Case 1: GetBoxScoreType = "boxscoreadvanced"
        Case 2: GetBoxScoreType = "boxscoreadvancedv2"
        Case 3: GetBoxScoreType = "boxscorefourfactors"
        Case 4: GetBoxScoreType = "boxscorefourfactorsv2"
        Case 5: GetBoxScoreType = "boxscoremisc"
        Case 6: GetBoxScoreType = "boxscoremiscv2"
        Case 7: GetBoxScoreType = "boxscoreplayertrackv2"
        Case 8: GetBoxScoreType = "boxscorescoring"
        Case 9: GetBoxScoreType = "boxscorescoringv2"
        Case 10: GetBoxScoreType = "boxscoresummaryv2"
        Case 11: GetBoxScoreType = "boxscoretraditionalv2"
        Case 12: GetBoxScoreType = "boxscoreusage"
        Case 13: GetBoxScoreType = "boxscoreusagev2"
    End Select
End Function

Private Function BoxScoreAPIAddress(GameID As String, StartPeriod As Period, EndPeriod As Period, BoxType As BoxScoreType) As String
    
    BoxScoreAPIAddress = "http://stats.nba.com/stats/" _
        & GetBoxScoreType(BoxType) _
        & "/?GameID=" & GameID _
        & "&StartPeriod=" & StartPeriod _
        & "&EndPeriod=" & EndPeriod _
        & "&StartRange=" & StartRange _
        & "&EndRange=" & EndRange _
        & "&RangeType=" & RangeType
End Function

Private Function CreateGameID(TwoDigitYear As Long, gamenumber As Long)
    Dim GameNumStr As String
    Dim i As Long
    
    GameNumStr = String(6 - Len(str(gamenumber)), "0") & gamenumber
    
    CreateGameID = "002" & TwoDigitYear & GameNumStr
End Function

'Private Sub downloadfile(link As String, FilePath As String)
'    Dim done
'
'    done = URLDownloadToFile(0, link, FilePath, 0, 0)
'
'     'Test.
'    If Not done = 0 Then
'        'msgBox "File not found!"
'    End If
'
'End Sub

Private Sub Class_initialize()
    Me.SeasonYear = Season2018
    pGameNumber = 1
    pStartPeriod = AllQuarters
    pEndPeriod = AllQuarters
    pBoxScoreType = Traditionalv2
End Sub




Public Sub GetPlayerstats(WorksheetToPrintTo As Worksheet, Optional vGameScope As ParGameScope = 0, Optional vPlayerExperience As ParPlayerExperience = 0, Optional vPlayerPosition As ParPlayerPosition = 0, Optional vStarterBench As ParStarterBench = 0, Optional vMeasureType As ParMeasureType = 0, Optional vPerMode As ParPerMode = 0, Optional vPlusMinus As ParPlusMinus = 0, Optional vPaceAdjust As ParPaceAdjust = 0, Optional vRank As ParRank = 0, Optional vSeason As Long = 2018, Optional vSeasonType As ParSeasonType = 0, Optional vOutcome As ParOutcome = 0, Optional vLocation As ParLocation = 0, Optional vMonth As ParMonth = 0, Optional vDateFrom As String = "", Optional vDateTo As String = "", Optional vOpposingTeam As String = "0", Optional vGameSegment As ParGameSegment = 0, Optional vPeriod As ParPeriod = 0, Optional vLastNGames As String = "0")
    
    Dim objXML As Object
    Dim url As String
    Dim JSON As String
    Dim dic As Dictionary
    
    url = BasePlayerStats & "GameScope=" & GetGameScope(vGameScope) & "&" & _
        "PlayerExperience=" & GetPlayerExperience(vPlayerExperience) & "&" & _
        "PlayerPosition=" & GetPlayerPosition(vPlayerPosition) & "&" & _
        "StarterBench=" & GetStarterBench(vStarterBench) & "&" & _
        "MeasureType=" & GetMeasureType(vMeasureType) & "&" & _
        "PerMode=" & GetPerMode(vPerMode) & "&" & _
        "PlusMinus=" & GetPlusMinus(vPlusMinus) & "&" & _
        "PaceAdjust=" & GetPaceAdjust(vPaceAdjust) & "&" & _
        "Rank=" & GetRank(vRank) & "&" & _
        "Season=" & GetSeason(vSeason) & "&" & _
        "SeasonType=" & GetSeasonType(vSeasonType) & "&" & _
        "Outcome=" & GetOutcome(vOutcome) & "&" & _
        "Location=" & GetLocation(vLocation) & "&" & _
        "Month=" & GetMonth(vMonth) & "&" & _
        "SeasonSegment=" & SeasonSegment & "&" & _
        "DateFrom=" & vDateFrom & "&" & _
        "DateTo=" & vDateTo & "&" & _
        "OpponentTeamID=" & GetTeamID(vOpposingTeam) & "&" & _
        "VsConference=" & VsConference & "&" & _
        "VsDivision=" & VsDivision & "&" & _
        "GameSegment=" & GetGameSegment(vGameSegment) & "&" & _
        "Period=" & GetPeriod(vPeriod) & "&" & _
        "LastNGames=" & vLastNGames
    
        'Debug.Print URL
    
    Set objXML = CreateObject("MSXML2.XMLHTTP")
    
    With objXML
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        JSON = .responseText
    End With
    
    Set dic = New Dictionary
    Set dic = ParseJson(JSON)
    
    Dim Name As String
    Dim Headers As Object
    Dim RowSet As Collection
    Name = dic("resultSets")(1)("name")
    'Debug.Print TypeName(dic("resultsSets"))
    Set Headers = dic("resultSets")(1)("headers")
    'Debug.Print TypeName(Headers)
    Set RowSet = dic("resultSets")(1)("rowSet")
    Dim Table() As Variant
    ReDim Table(1 To RowSet.Count, 1 To RowSet(1).Count)
    Dim i As Long
    Dim j As Long
    For i = 1 To RowSet.Count
        For j = 1 To RowSet(1).Count
            Table(i, j) = RowSet(i)(j)
        Next j
    Next i
    Dim k As Long
    For k = 1 To Headers.Count
        WorksheetToPrintTo.Range("A1").Offset(0, k - 1).Value = Headers(k)
    Next k
    
    Range(WorksheetToPrintTo.Range("A2"), WorksheetToPrintTo.Range("A2").Offset(RowSet.Count - 1, RowSet(1).Count - 1)).Value = Table
    
End Sub


Public Function GetTeamVsOpponent(WorksheetToPrintTo As Worksheet, Optional vTeamID As String = "", Optional vMeasureType As ParMeasureType = 0, Optional vPerMode As ParPerMode = 0, Optional vPlusMinus As ParPlusMinus = 0, Optional vPaceAdjust As ParPaceAdjust = 0, Optional vRank As ParRank = 0, Optional vSeason As Long = 2018, Optional vSeasonType As ParSeasonType = 0, Optional vOutcome As ParOutcome = 0, Optional vLocation As ParLocation = 0, Optional vMonth As ParMonth = 0, Optional vDateFrom As String = "", Optional vDateTo As String = "", Optional vOpposingTeam As String = "0", Optional vGameSegment As ParGameSegment = 0, Optional vPeriod As ParPeriod = 0, Optional vLastNGames As String = "0") As String
    
    Dim objXML As Object
    Dim url As String
    Dim JSON As String
    Dim dic As Dictionary
    
    url = BaseTeamVsOpponent & "TeamID=" & GetTeamID(vTeamID) & "&" & "TeamID=" & GetTeamID(vTeamID) & "&" & _
        "MeasureType=" & GetMeasureType(vMeasureType) & "&" & _
        "PerMode=" & GetPerMode(vPerMode) & "&" & _
        "PlusMinus=" & GetPlusMinus(vPlusMinus) & "&" & _
        "PaceAdjust=" & GetPaceAdjust(vPaceAdjust) & "&" & _
        "Rank=" & GetRank(vRank) & "&" & _
        "Season=" & GetSeason(vSeason) & "&" & _
        "SeasonType=" & GetSeasonType(vSeasonType) & "&" & _
        "Outcome=" & GetOutcome(vOutcome) & "&" & _
        "Location=" & GetLocation(vLocation) & "&" & _
        "Month=" & GetMonth(vMonth) & "&" & _
        "SeasonSegment=" & SeasonSegment & "&" & _
        "DateFrom=" & vDateFrom & "&" & _
        "DateTo=" & vDateTo & "&" & _
        "OpponentTeamID=" & GetTeamID(vOpposingTeam) & "&" & _
        "VsConference=" & VsConference & "&" & _
        "VsDivision=" & VsDivision & "&" & _
        "GameSegment=" & GetGameSegment(vGameSegment) & "&" & _
        "Period=" & GetPeriod(vPeriod) & "&" & _
        "LastNGames=" & vLastNGames
    
        Debug.Print url
    
    Set objXML = CreateObject("MSXML2.XMLHTTP")
    
    With objXML
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        JSON = .responseText
    End With
    'msgBox JSON
    Set dic = New Dictionary
    Set dic = ParseJson(JSON)
    
    Dim Name As String
    Dim Headers As Object
    Dim RowSet As Collection
    Name = dic("resultSets")(1)("name")
    'Debug.Print TypeName(dic("resultsSets"))
    Set Headers = dic("resultSets")(1)("headers")
    Debug.Print TypeName(Headers)
    Set RowSet = dic("resultSets")(1)("rowSet")
    Dim Table() As Variant
    ReDim Table(1 To RowSet.Count, 1 To RowSet(1).Count)
    Dim i As Long
    Dim j As Long
    For i = 1 To RowSet.Count
        For j = 1 To RowSet(1).Count
            Table(i, j) = RowSet(i)(j)
        Next j
    Next i
    Dim k As Long
    For k = 1 To Headers.Count
        WorksheetToPrintTo.Range("A1").Offset(0, k - 1).Value = Headers(k)
    Next k
    
    Range(WorksheetToPrintTo.Range("A2"), WorksheetToPrintTo.Range("A2").Offset(RowSet.Count - 1, RowSet(1).Count - 1)).Value = Table
    
End Function


Public Function GetTeamStats(WorksheetToPrintTo As Worksheet, Optional vTeamID As String, Optional vMeasureType As ParMeasureType = 0, Optional vPerMode As ParPerMode = 0, Optional vPlusMinus As ParPlusMinus = 0, Optional vPaceAdjust As ParPaceAdjust = 0, Optional vRank As ParRank = 0, Optional vSeason As Long = 2018, Optional vSeasonType As ParSeasonType = 0, Optional vOutcome As ParOutcome = 0, Optional vLocation As ParLocation = 0, Optional vMonth As ParMonth = 0, Optional vDateFrom As String = "", Optional vDateTo As String = "", Optional vOpposingTeam As String = "0", Optional vGameSegment As ParGameSegment = 0, Optional vPeriod As ParPeriod = 0, Optional vLastNGames As String = "0") As String
    
    Dim objXML As Object
    Dim url As String
    Dim JSON As String
    Dim dic As Dictionary
    
    url = Base & "leaguedashteamstats/?" & _
        "MeasureType=" & GetMeasureType(vMeasureType) & "&" & _
        "PerMode=" & GetPerMode(vPerMode) & "&" & _
        "PlusMinus=" & GetPlusMinus(vPlusMinus) & "&" & _
        "PaceAdjust=" & GetPaceAdjust(vPaceAdjust) & "&" & _
        "Rank=" & GetRank(vRank) & "&" & _
        "Season=" & GetSeason(vSeason) & "&" & _
        "SeasonType=" & GetSeasonType(vSeasonType) & "&" & _
        "Outcome=" & GetOutcome(vOutcome) & "&" & _
        "Location=" & GetLocation(vLocation) & "&" & _
        "Month=" & GetMonth(vMonth) & "&" & _
        "SeasonSegment=" & SeasonSegment & "&" & _
        "DateFrom=" & vDateFrom & "&" & _
        "DateTo=" & vDateTo & "&" & _
        "OpponentTeamID=" & GetTeamID(vOpposingTeam) & "&" & _
        "VsConference=" & VsConference & "&" & _
        "VsDivision=" & VsDivision & "&" & _
        "GameSegment=" & GetGameSegment(vGameSegment) & "&" & _
        "Period=" & GetPeriod(vPeriod) & "&" & _
        "LastNGames=" & vLastNGames
    
        Debug.Print url
    
    Set objXML = CreateObject("MSXML2.XMLHTTP")
    
    With objXML
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        JSON = .responseText
    End With
    'msgBox JSON
    Set dic = New Dictionary
    Set dic = ParseJson(JSON)
    
    Dim Name As String
    Dim Headers As Object
    Dim RowSet As Collection
    Name = dic("resultSets")(1)("name")
    'Debug.Print TypeName(dic("resultsSets"))
    Set Headers = dic("resultSets")(1)("headers")
    Debug.Print TypeName(Headers)
    Set RowSet = dic("resultSets")(1)("rowSet")
    Dim Table() As Variant
    ReDim Table(1 To RowSet.Count, 1 To RowSet(1).Count)
    Dim i As Long
    Dim j As Long
    For i = 1 To RowSet.Count
        For j = 1 To RowSet(1).Count
            Table(i, j) = RowSet(i)(j)
        Next j
    Next i
    Dim k As Long
    For k = 1 To Headers.Count
        WorksheetToPrintTo.Range("A1").Offset(0, k - 1).Value = Headers(k)
    Next k
    
    Range(WorksheetToPrintTo.Range("A2"), WorksheetToPrintTo.Range("A2").Offset(RowSet.Count - 1, RowSet(1).Count - 1)).Value = Table
    
End Function

Public Function GetTeamOnOff(WorksheetToPrintTo As Worksheet, Optional vTeamID As String, Optional vMeasureType As ParMeasureType = 0, Optional vPerMode As ParPerMode = 0, Optional vPlusMinus As ParPlusMinus = 0, Optional vPaceAdjust As ParPaceAdjust = 0, Optional vRank As ParRank = 0, Optional vSeason As Long = 2018, Optional vSeasonType As ParSeasonType = 0, Optional vOutcome As ParOutcome = 0, Optional vLocation As ParLocation = 0, Optional vMonth As ParMonth = 0, Optional vDateFrom As String = "", Optional vDateTo As String = "", Optional vOpposingTeam As String = "0", Optional vGameSegment As ParGameSegment = 0, Optional vPeriod As ParPeriod = 0, Optional vLastNGames As String = "0") As String
    
    Dim objXML As Object
    Dim url As String
    Dim JSON As String
    Dim dic As Dictionary
    
    url = Base & "teamplayeronoffdetails/?" & "&" & "TeamID=" & GetTeamID(vTeamID) & "&" & _
        "MeasureType=" & GetMeasureType(vMeasureType) & "&" & _
        "PerMode=" & GetPerMode(vPerMode) & "&" & _
        "PlusMinus=" & GetPlusMinus(vPlusMinus) & "&" & _
        "PaceAdjust=" & GetPaceAdjust(vPaceAdjust) & "&" & _
        "Rank=" & GetRank(vRank) & "&" & _
        "Season=" & GetSeason(vSeason) & "&" & _
        "SeasonType=" & GetSeasonType(vSeasonType) & "&" & _
        "Outcome=" & GetOutcome(vOutcome) & "&" & _
        "Location=" & GetLocation(vLocation) & "&" & _
        "Month=" & GetMonth(vMonth) & "&" & _
        "SeasonSegment=" & SeasonSegment & "&" & _
        "DateFrom=" & vDateFrom & "&" & _
        "DateTo=" & vDateTo & "&" & _
        "OpponentTeamID=" & GetTeamID(vOpposingTeam) & "&" & _
        "VsConference=" & VsConference & "&" & _
        "VsDivision=" & VsDivision & "&" & _
        "GameSegment=" & GetGameSegment(vGameSegment) & "&" & _
        "Period=" & GetPeriod(vPeriod) & "&" & _
        "LastNGames=" & vLastNGames
    
        Debug.Print url
    
    Set objXML = CreateObject("MSXML2.XMLHTTP")
    
    With objXML
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        JSON = .responseText
    End With
    'msgBox JSON
    Set dic = New Dictionary
    Set dic = ParseJson(JSON)
    
    Dim Name As String
    Dim Headers As Object
    Dim RowSet As Collection
    Name = dic("resultSets")(1)("name")
    'Debug.Print TypeName(dic("resultsSets"))
    Set Headers = dic("resultSets")(1)("headers")
    Debug.Print TypeName(Headers)
    Set RowSet = dic("resultSets")(1)("rowSet")
    Dim Table() As Variant
    ReDim Table(1 To RowSet.Count, 1 To RowSet(1).Count)
    Dim i As Long
    Dim j As Long
    For i = 1 To RowSet.Count
        For j = 1 To RowSet(1).Count
            Table(i, j) = RowSet(i)(j)
        Next j
    Next i
    Dim k As Long
    For k = 1 To Headers.Count
        WorksheetToPrintTo.Range("A1").Offset(0, k - 1).Value = Headers(k)
    Next k
    
    Range(WorksheetToPrintTo.Range("A2"), WorksheetToPrintTo.Range("A2").Offset(RowSet.Count - 1, RowSet(1).Count - 1)).Value = Table
    
End Function

Public Sub GetBoxScore(WorksheetToPrintTo As Worksheet)
    
    Dim objXML As Object
    Dim url As String
    Dim JSON As String
    Dim dic As Dictionary
    
    
    url = BoxScoreAPIAddress(CreateGameID(pSeasonYear, pGameNumber), pStartPeriod, pEndPeriod, pBoxScoreType)
    
    Set objXML = CreateObject("MSXML2.XMLHTTP")
    
    With objXML
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        JSON = .responseText
    End With
    'msgBox JSON
    Set dic = New Dictionary
    Set dic = ParseJson(JSON)
    
    Dim Name As String
    Dim Headers As Object
    Dim RowSet As Collection
    Name = dic("resultSets")(1)("name")
    'Debug.Print TypeName(dic("resultsSets"))
    Set Headers = dic("resultSets")(1)("headers")
    'Debug.Print TypeName(Headers)
    Set RowSet = dic("resultSets")(1)("rowSet")
    Dim Table() As Variant
    ReDim Table(1 To RowSet.Count, 1 To RowSet(1).Count)
    Dim i As Long
    Dim j As Long
    For i = 1 To RowSet.Count
        For j = 1 To RowSet(1).Count
            Table(i, j) = "'" & RowSet(i)(j)
        Next j
    Next i
    Dim k As Long
    For k = 1 To Headers.Count
        WorksheetToPrintTo.Range("A1").Offset(1, k - 1).Value = Headers(k)
    Next k
    
    Dim vOffset As Long
    vOffset = WorksheetToPrintTo.Range("A999999").End(xlUp).Row - 1
    
    Range(WorksheetToPrintTo.Range("A2").Offset(vOffset, 0), WorksheetToPrintTo.Range("A2").Offset(vOffset, 0).Offset(RowSet.Count - 1, RowSet(1).Count - 1)).Value = Table
    
End Sub



