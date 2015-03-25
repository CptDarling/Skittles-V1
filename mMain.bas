Attribute VB_Name = "mMain"
Option Explicit

'Public Options As cOptions
Public fMain As frmMain
Public fStatistics As frmStatistics
Public fTransactions As frmTransactions
Public ShuttingDown As Boolean
Public Sub Main()
'    Set Options = New cOptions
    Set fMain = New frmMain
    fMain.Show
End Sub

Public Sub ExitApp()
    On Error Resume Next
    ShuttingDown = True
    
    Unload fTransactions
    Set fTransactions = Nothing
    
    Unload fStatistics
    Set fStatistics = Nothing
    
    Unload fMain
    Set fMain = Nothing
    
'    Set Options = Nothing
    
    End
End Sub

Public Sub ShowAbout()
    MsgBox App.Title & vbCrLf & "Version " & App.Major & "." & App.Minor & " revision " & App.Revision & vbCrLf & App.LegalCopyright, vbInformation, "About " & App.Title & "..."
End Sub

Public Function Alert(Message As Variant, Optional Style As VbMsgBoxStyle = vbInformation, Optional Title As String) As VbMsgBoxResult
    Dim sMessage As String
    
    If IsMissing(Title) Then Title = App.Title
    
    If IsObject(Message) Then
        sMessage = Err.Number & vbCrLf & Err.Source & vbCrLf & Err.Description
    Else
        sMessage = Message
    End If
    
    Alert = MsgBox(sMessage, Style, Title)

End Function
