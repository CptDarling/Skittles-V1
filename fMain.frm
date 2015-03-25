VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Skittles Database"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9735
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5025
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16642
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuMatchStatistics 
         Caption         =   "Match statistics..."
      End
      Begin VB.Menu mnuTransactions 
         Caption         =   "&Transactions..."
      End
      Begin VB.Menu mnuAddPoints 
         Caption         =   "Enter &Points..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalcStats 
         Caption         =   "Calculate &statistics..."
      End
      Begin VB.Menu EraseStats 
         Caption         =   "&Erase all statistics"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Skittles Database..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EraseStats_Click()
    de.EraseStatistics
End Sub


Private Sub MDIForm_Load()
    GetWindowPosition Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    PutWindowPosition Me, True
    If Not ShuttingDown Then ExitApp
End Sub


Private Sub mnuAbout_Click()
    ShowAbout
End Sub

Private Sub mnuAddPoints_Click()
    Dim sMatchID As String
    sMatchID = InputBox("Enter the Match ID.", "Enter Match Points")
    If Not IsNumeric(sMatchID) Then Exit Sub
    AddPoints CInt(sMatchID)
End Sub

Private Sub mnuCalcStats_Click()
    Dim sTmp As String
    Dim sTitle As String
    
    sTitle = "Calculate Match Statistics"
    
    sTmp = InputBox("Enter the ID number of the season to be calculated.", sTitle, RegGet(REG_SETTINGS, "Season", 2))
    If Not IsNumeric(sTmp) Then Exit Sub
    
    RegPut REG_SETTINGS, "Season", sTmp, REG_SZ
    Statistics CLng(sTmp)
    MsgBox "Calculation of match statistics is complete.", vbInformation, sTitle

End Sub

Private Sub mnuExit_Click()
    ExitApp
End Sub


Private Sub mnuMatchStatistics_Click()
    If fStatistics Is Nothing Then Set fStatistics = New frmStatistics
    Me.mnuMatchStatistics.Checked = Not Me.mnuMatchStatistics.Checked
    If Me.mnuMatchStatistics.Checked Then
        fStatistics.Show
    Else
        fStatistics.Hide
    End If
End Sub


Private Sub mnuTransactions_Click()
    If fTransactions Is Nothing Then Set fTransactions = New frmTransactions
    Me.mnuTransactions.Checked = Not Me.mnuTransactions.Checked
    If Me.mnuTransactions.Checked Then
        fTransactions.Show
    Else
        fTransactions.Hide
    End If
End Sub


