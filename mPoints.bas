Attribute VB_Name = "mPoints"
Option Explicit


Public Sub AddPoints(MatchID As Integer)
    Dim iLegs As Integer
    Dim i As Integer
    Dim sPoints As String
    With de.rstblPoints
        .Open
        de.MatchLegs MatchID
        iLegs = de.rsMatchLegs.Fields("Legs").Value
        de.rsMatchLegs.Close
        For i = 1 To iLegs
            sPoints = InputBox("Match " & MatchID & " points for leg " & i & "", "Input Points", 0)
            If Not IsNumeric(sPoints) Then Exit For
            .AddNew Array("MatchID", "Leg", "Points"), Array(MatchID, i, sPoints)
        Next i
        .Close
    End With
End Sub
