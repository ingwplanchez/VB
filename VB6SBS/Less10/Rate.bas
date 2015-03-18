Attribute VB_Name = "Module1"
Public Wins
Public Spins
Function Rate(Hits, Attempts) As String
    Percent = Hits / Attempts
    Rate = Format(Percent, "0.0%")
End Function
