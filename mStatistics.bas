Attribute VB_Name = "mStatistics"
Option Explicit


Public Enum sdbStatAge
    sdbCurrent = 1
    sdbComparison = 2
End Enum


Public Enum statDirection
    Ascending = 1
    Descending = 2
End Enum

Private StatisticID As Long

Private Sub CalculateStats(SeasonID As Integer, MatchDate As Date, CompareDate As Date)
    
    MatchesPlayedByPlayer SeasonID, MatchDate, sdbCurrent
    MatchesPlayedByPlayer SeasonID, CompareDate, sdbComparison
    
    TotalScore SeasonID, MatchDate, sdbCurrent
    TotalScore SeasonID, CompareDate, sdbComparison
    
    TotalMisses SeasonID, MatchDate, sdbCurrent
    TotalMisses SeasonID, CompareDate, sdbComparison
    
    ScorePerMatch SeasonID, MatchDate, sdbCurrent
    ScorePerMatch SeasonID, CompareDate, sdbComparison
    
    MissesPerMatch SeasonID, MatchDate, sdbCurrent
    MissesPerMatch SeasonID, CompareDate, sdbComparison
    
    Stacks SeasonID, MatchDate, sdbCurrent
    Stacks SeasonID, CompareDate, sdbComparison
    
    HighScores SeasonID, MatchDate, sdbCurrent
    HighScores SeasonID, CompareDate, sdbComparison
    
    TotalFines SeasonID, MatchDate, sdbCurrent
    TotalFines SeasonID, CompareDate, sdbComparison
    
    FinesPerMatch SeasonID, MatchDate, sdbCurrent
    FinesPerMatch SeasonID, CompareDate, sdbComparison
    
    'System("StatisticsCurrentMatchDate") = MatchDate
    'System("StatisticsPreviousMatchDate") = CompareDate
    
End Sub

Private Sub MatchesPlayedByPlayer(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
    sPrefix = GetPrefix(StatAge)
    
    de.MatchesPlayedByPlayer MatchDate, SeasonID, "%"
    With de.rsMatchesPlayedByPlayer
        Do Until .EOF
            de.tblStatistics !MemberID, 1, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = !CountOfMemberID
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 1, Descending, SeasonID, sPrefix
    
End Sub

Private Sub TotalScore(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
    sPrefix = GetPrefix(StatAge)
    
    de.TotalScore MatchDate, SeasonID
    With de.rsTotalScore
        Do Until .EOF
            de.tblStatistics !MemberID, 2, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = !SumOfScore
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 2, Descending, SeasonID, sPrefix
    
End Sub


Private Sub ScorePerMatch(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
'    Dim iLegs As Integer
'    iLegs = GetLegs(MatchDate, SeasonID)
    sPrefix = GetPrefix(StatAge)
    de.TotalScore MatchDate, SeasonID
    With de.rsTotalScore
        Do Until .EOF
            de.MatchesPlayedByPlayer MatchDate, SeasonID, !MemberID
            de.tblStatistics !MemberID, 3, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = !SumOfScore / de.rsMatchesPlayedByPlayer!MatchesPlayedByLegs
If !MemberID = "VG" And MatchDate = "24/11/2004" Then
Debug.Print
End If
            de.rstblStatistics.Fields(sPrefix & "Value02").Value = !SumOfScore / de.rsMatchesPlayedByPlayer!LegsPlayed
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            de.rsMatchesPlayedByPlayer.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 3, Descending, SeasonID, sPrefix
    
End Sub



Private Sub TotalMisses(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
    sPrefix = GetPrefix(StatAge)
    
    de.TotalMisses MatchDate, SeasonID
    With de.rsTotalMisses
        Do Until .EOF
            de.tblStatistics !MemberID, 4, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = !SumOfMisses - !SumOfFines
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 4, Ascending, SeasonID, sPrefix
    
End Sub

Private Sub TotalFines(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
    sPrefix = GetPrefix(StatAge)
    
    de.TotalMisses MatchDate, SeasonID
    With de.rsTotalMisses
        Do Until .EOF
            de.tblStatistics !MemberID, 8, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = !SumOfFines
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 8, Ascending, SeasonID, sPrefix
    
End Sub


Private Sub Stacks(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String, iPos As Long, dLast As Double, i As Integer
    sPrefix = GetPrefix(StatAge)
    
    de.Stacks MatchDate, SeasonID, "%"
    With de.rsStacks
        Do Until .EOF
            de.tblStatistics !MemberID, 6, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = !CountOfStacks
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    If sPrefix = "Comparison" Then If MatchDate = 0 Then GoTo Skip_
    
    de.NoStacks StatisticID
    With de.rsNoStacks
        Do Until .EOF
            .Fields(sPrefix & "Date").Value = MatchDate
            .Fields(sPrefix & "Value01").Value = 0
            .Update
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
Skip_:
    ReOrder 6, Descending, SeasonID, sPrefix
    
End Sub


Private Sub HighScores(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
    sPrefix = GetPrefix(StatAge)
    
    de.HighScores MatchDate, SeasonID, "%"
    With de.rsHighScores
        Do Until .EOF
            de.tblStatistics !MemberID, 7, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = !MaxOfScore
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 7, Descending, SeasonID, sPrefix
    
End Sub



Private Sub MissesPerMatch(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
'    Dim iLegs As Integer
'    iLegs = GetLegs(MatchDate, SeasonID)
    sPrefix = GetPrefix(StatAge)
    
    de.TotalMisses MatchDate, SeasonID
    With de.rsTotalMisses
        Do Until .EOF
            de.MatchesPlayedByPlayer MatchDate, SeasonID, !MemberID
            de.tblStatistics !MemberID, 5, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = (!SumOfMisses - !SumOfFines) / de.rsMatchesPlayedByPlayer!MatchesPlayedByLegs
            de.rstblStatistics.Fields(sPrefix & "Value02").Value = (!SumOfMisses - !SumOfFines) / de.rsMatchesPlayedByPlayer!LegsPlayed
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            de.rsMatchesPlayedByPlayer.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 5, Ascending, SeasonID, sPrefix

End Sub


Private Sub FinesPerMatch(SeasonID As Integer, MatchDate As Date, StatAge As sdbStatAge)
    Dim sPrefix As String
'    Dim iLegs As Integer
'    iLegs = GetLegs(MatchDate, SeasonID)
    sPrefix = GetPrefix(StatAge)
    
    de.TotalMisses MatchDate, SeasonID
    With de.rsTotalMisses
        Do Until .EOF
            de.MatchesPlayedByPlayer MatchDate, SeasonID, !MemberID
            de.tblStatistics !MemberID, 9, StatisticID
            de.rstblStatistics.Fields(sPrefix & "Date").Value = MatchDate
            de.rstblStatistics.Fields(sPrefix & "Value01").Value = (!SumOfFines) / de.rsMatchesPlayedByPlayer!MatchesPlayedByLegs
            de.rstblStatistics.Fields(sPrefix & "Value02").Value = (!SumOfFines) / de.rsMatchesPlayedByPlayer!LegsPlayed
            de.rstblStatistics.Update
            de.rstblStatistics.Close
            de.rsMatchesPlayedByPlayer.Close
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    ReOrder 9, Ascending, SeasonID, sPrefix

End Sub



Private Sub CreateRows(StatisticID As Long, SeasonID As Integer, LatestMatchDate As Date)
    Dim sPrefix As String
    Dim i As Integer
    Dim iStatTypes As Integer
    sPrefix = GetPrefix(sdbCurrent)
    iStatTypes = CountOfStatTypes
    de.MatchesPlayedByPlayer LatestMatchDate, SeasonID, "%"
    With de.rsMatchesPlayedByPlayer
        de.tblStatistics 1, "", StatisticID
        Do Until .EOF
            For i = 1 To iStatTypes
                de.rstblStatistics.AddNew Array("StatisticID", "StatisticGroup", "MemberID", "SeasonID"), Array(StatisticID, i, !MemberID, SeasonID)
            Next i
            .MoveNext
            DoEvents
        Loop
        de.rstblStatistics.Close
        .Close
    End With
    
End Sub


Private Function GetPrefix(StatAge As sdbStatAge) As String
    Select Case StatAge
        Case sdbStatAge.sdbComparison
            GetPrefix = "Comparison"
        Case sdbStatAge.sdbCurrent
            GetPrefix = "Current"
    End Select
End Function

Private Sub CalculateChanges(FieldName As String)
    Dim dblCurrent As Variant, dblComparison As Variant, dblResult As Double, sText As String
    With de.rstblStatisticsWrite
        .Open
        On Error Resume Next
        Do Until .EOF
            dblCurrent = 0
            dblComparison = 0
            dblCurrent = .Fields("Current" & FieldName).Value
            dblComparison = .Fields("Comparison" & FieldName).Value
            dblResult = dblCurrent - dblComparison
            If Not IsNull(dblCurrent) And Not IsNull(dblComparison) Then
                Select Case dblResult
                    Case 0
                        sText = "="
                    Case dblCurrent
                        sText = "-"
                    Case Is > 0
                        If FieldName = "Position" Then
                            sText = "d" & dblComparison
                        Else
                            sText = "u" & Format(dblComparison, "Fixed")
                        End If
                    Case Is < 0
                        If FieldName = "Position" Then
                            sText = "u" & dblComparison
                        Else
                            sText = "d" & Format(dblComparison, "Fixed")
                        End If
                End Select
            Else
                sText = "-"
            End If
            .Fields("Change" & FieldName).Value = sText
            .Update
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
End Sub

Public Sub ReOrder(StatisticGroup As Integer, Direction As statDirection, SeasonID As Integer, sPrefix As String)
    Dim iPos As Long
    Dim dLast As Double
    Dim i As Integer
    
    de.ReorderStatistics "%", SeasonID, StatisticGroup, StatisticID
    iPos = 1: i = 0: dLast = 0
    With de.rsReorderStatistics
        .Sort = sPrefix & "Value01 " & Choose(Direction, "ASC", "DESC") & ", MemberID"
        Do Until .EOF
            i = i + 1
            If dLast <> .Fields(sPrefix & "Value01").Value Then iPos = i
            If Not IsNull(.Fields(sPrefix & "Value01").Value) Then
                dLast = .Fields(sPrefix & "Value01").Value
                .Update sPrefix & "Position", iPos
            Else
                dLast = 0
            End If
            .MoveNext
            DoEvents
        Loop
        .Close
    End With

End Sub

Public Sub Statistics(SeasonID As Integer)
    Dim dCurrent As Date, dLast As Date
    Dim id As Long
    Dim rec As Long, recs As Long

    de.EraseStatistics
    de.MatchDates SeasonID
    
    With de.rsMatchDates
        recs = .RecordCount
        id = 1
        Do Until .EOF
            dCurrent = !Date
            Status "Comparing " & dCurrent & " to " & dLast & "..."
            StatisticID = id
            'If Not Exists(dCurrent) Then
                CreateRows id, SeasonID, dCurrent
                CalculateStats SeasonID, dCurrent, dLast
            'End If
            dLast = dCurrent
            id = id + 1
            rec = rec + 1
            .MoveNext
            DoEvents
        Loop
        .Close
    End With
    
    Status "Calculating changes in Value01..."
    CalculateChanges "Value01"
    Status "Calculating changes in Value02..."
    CalculateChanges "Value02"
    Status "Calculating changes in position..."
    CalculateChanges "Position"
    
    Status

End Sub

Private Function Exists(StatisticDate As Date) As Boolean
    de.StatisticDates Format(StatisticDate, "yymmdd")
    Exists = Not de.rsStatisticDates.BOF
    de.rsStatisticDates.Close
End Function

Private Function CountOfStatTypes() As Integer
    On Error Resume Next
    With de.rsCountOfStatTypes
        .Open
        CountOfStatTypes = !CountOfStatTypes
        .Close
    End With
End Function

'Public Function GetLegs(MatchDate As Date, SeasonID As Integer) As Integer
'    On Error Resume Next
'    de.TotalLegs MatchDate, SeasonID
'    GetLegs = de.rsTotalLegs("CountOfLegs").Value
'    de.rsTotalLegs.Close
'End Function
