Attribute VB_Name = "Module1"
Option Base 1         'set array base to 1
Public Scoreboard(2, 9) As Variant
Public Inning As Integer

Sub AddUpScores()
'AddUpScores is a public procedure that totals and
'displays the runs in the Scoreboard array.

    For i% = 1 To 9   'use loop to add scores
        AwayScore% = AwayScore% + Scoreboard(1, i%)
        HomeScore% = HomeScore% + Scoreboard(2, i%)
    Next i%           'then display scores in box
    Form1.CurrentX = 5000
    Form1.CurrentY = 1050
    Form1.Print AwayScore%
    Form1.CurrentX = 5000
    Form1.CurrentY = 1400
    Form1.Print HomeScore%
End Sub


