Attribute VB_Name = "Scraping5"
Sub getbasicandadvancedbox()
    Call GetAllBasicBoxScores
    Call GetAllAdvancedBoxScores
    
End Sub
Sub GetAllBasicBoxScores()

    Application.ScreenUpdating = False
    
    Dim Tracker As Worksheet
    Workbooks.Add
    Set Tracker = ActiveSheet
    With Tracker
        .Name = "Tracker"
        .Range("A1") = "Game_ID"
        .Range("B1") = "Date"
        .Range("C1") = "Team1"
        .Range("D1") = "Team2"
    End With
    
    Dim objNBA As NBAStats
    Set objNBA = New NBAStats
    Dim i As Long
    On Error GoTo FInished
    For i = 1 To 1000
        objNBA.gamenumber = i
        Call objNBA.GetBoxScore(Worksheets.Add)
        ActiveSheet.Name = ActiveSheet.Range("A2").Value
        
        With Tracker
            .Range("A10000").End(xlUp).Offset(1, 0) = ActiveSheet.Range("A2")
            .Range("B1") = "?"
            .Range("C90000").End(xlUp).Offset(1, 0) = ActiveSheet.Range("C2")
            .Range("D200000").End(xlUp).Offset(1, 0) = ActiveSheet.Range("c2").End(xlDown)
        End With
    Next i
    
    Application.ScreenUpdating = True
    
FInished:
    
End Sub


Sub GetAllAdvancedBoxScores()

    Application.ScreenUpdating = False
    
    Dim Tracker As Worksheet
    Workbooks.Add
    Set Tracker = ActiveSheet
    With Tracker
        .Name = "Tracker"
        .Range("A1") = "Game_ID"
        .Range("B1") = "Date"
        .Range("C1") = "Team1"
        .Range("D1") = "Team2"
    End With
    
    Dim objNBA As NBAStats
    Set objNBA = New NBAStats
    objNBA.BoxScoreCategory = Advancedv2
    Dim i As Long
    On Error GoTo FInished
    For i = 1 To 1000
        objNBA.gamenumber = i
        Call objNBA.GetBoxScore(Worksheets.Add)
        ActiveSheet.Name = ActiveSheet.Range("A2").Value
        
        With Tracker
            .Range("A10000").End(xlUp).Offset(1, 0) = ActiveSheet.Range("A2")
            .Range("B1") = "?"
            .Range("C90000").End(xlUp).Offset(1, 0) = ActiveSheet.Range("C2")
            .Range("D200000").End(xlUp).Offset(1, 0) = ActiveSheet.Range("c2").End(xlDown)
        End With
    Next i
    
    Application.ScreenUpdating = True
    
FInished:
    
End Sub
