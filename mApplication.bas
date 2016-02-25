Attribute VB_Name = "mApplication"
Option Explicit

Public Sub InitMatch(MatchID As Integer, Players As Collection)
    Dim sShort As Variant
    Dim i As Integer
    Dim sScore As String
    Dim sMisses As String
    Dim ii As Integer
    Dim iLegs As Integer
    de.MatchLegs MatchID
    iLegs = de.rsMatchLegs("Legs").Value
    de.rsMatchLegs.Close
    With de.rstblMatchData
        .Open
        For Each sShort In Players
            ii = ii + 1
            sShort = UCase(sShort)
            For i = 1 To iLegs
                sScore = InputBox(sShort & ", leg = " & i & " score?", "Score", "")
                sMisses = InputBox(sShort & ", leg = " & i & " misses?", "Misses", "0")
                
                
                .AddNew Array("MatchID", "MemberID", "PlayerPosition", "Leg", "Score", "Misses"), Array(MatchID, sShort, ii, i, Val(sScore), Val(sMisses))
            Next i
        Next sShort
        .Close
    End With
End Sub



Public Property Get System(ByVal Property As String) As String
    On Error GoTo Error_
    de.tblSystem CStr(Property)
    With de.rstblSystem
        If Not .BOF And Not .EOF Then System = !Value
        .Close
    End With

Exit_:
    On Error Resume Next
    de.rstblSystem.Close
    Exit Sub
    
Error_:
    System = "#Error: " & Err.Number & " " & Err.Description
    Resume Exit_

End Property

Public Property Let System(ByVal Property As String, ByVal vNewValue As String)
    On Error Resume Next
    de.tblSystem CStr(Property)
    With de.rstblSystem
        If .BOF Or .EOF Then
            .AddNew Array("Property", "Value"), Array(Property, vNewValue)
        Else
            !Value = vNewValue
            .Update
        End If
        .Close
    End With
End Property

Public Sub Status(Optional Message As String)
    fMain.stbStatus.Panels(1).Text = Message
End Sub
