VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Match Statistics"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   9030
   Begin VB.CommandButton cmdInputGame 
      Caption         =   "Input Match Details..."
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   2760
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstLegs 
      Height          =   2655
      Left            =   3120
      TabIndex        =   16
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Legs"
         Object.Width           =   1191
      EndProperty
   End
   Begin VB.ComboBox cboMatch 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3855
   End
   Begin VB.TextBox txtMisses 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   8160
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMisses 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   7560
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMisses 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6960
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMisses 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6360
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMisses 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMisses 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMisses 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   8160
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   7560
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6960
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin MSComctlLib.ListView lstPlayers 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Players"
         Object.Width           =   5424
      EndProperty
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInputGame_Click()
    Dim sMatch As String
    Dim sTitle As String
    Dim sTmp As String
    Dim colPlayers As New Collection
    Dim i As Integer
    
    sTitle = "Input Match Data"
    
    sMatch = InputBox("Enter match ID.", sTitle)
    If sMatch = "" Or Not IsNumeric(sMatch) Then GoTo Exit_
    
    sTmp = "X"
    Do Until sTmp = "" Or LCase(sTmp) = "end"
        i = i + 1
        sTmp = InputBox("Enter player " & i & " ID. Enter 'end' to finish entering the list of players.", sTitle)
        If sTmp <> "" And LCase(sTmp) <> "end" Then colPlayers.Add UCase(sTmp)
    Loop
    If sTmp = "" Then GoTo Exit_
        
    InitMatch CInt(sMatch), colPlayers

Exit_:
    Set colPlayers = Nothing

End Sub


Private Sub Form_Load()
    GetWindowPosition Me
    LoadLegList
    LoadMatchList
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = ShuttingDown
End Sub


Private Sub Form_Resize()
    Me.lstLegs.ColumnHeaders(1).Width = Me.lstLegs.Width - 60
    Me.lstPlayers.ColumnHeaders(1).Width = Me.lstPlayers.Width - 60
End Sub


Private Sub Form_Unload(Cancel As Integer)
    PutWindowPosition Me, True
    fMain.mnuMatchStatistics.Checked = False
End Sub



Public Sub LoadLegList()
    Dim i As Integer
    With Me.lstLegs.ListItems
        For i = 1 To 7
            .Add , , i
        Next i
        .Item(1).Selected = True
    End With
End Sub

Private Sub lstLegs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    MsgBox ColumnHeader.Width & "/" & Me.lstLegs.Width
End Sub


Private Sub lstPlayers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    MsgBox ColumnHeader.Width & "/" & Me.lstPlayers.Width
End Sub



Public Sub LoadMatchList()
    de.rstblMatches.Open
    With de.rstblMatches
        Do Until .EOF
            Me.cboMatch.AddItem !LookupSeason!Season & ", " & !Date & ", " & !Verses
            .MoveNext
        Loop
    End With

Exit_:
    On Error Resume Next
    de.rstblMatches.Close
    Exit Sub

Error_:
    Alert Err, vbCritical, Me.Name & ".LoadMatchList"
    Resume Exit_
    
End Sub
