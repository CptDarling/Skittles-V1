VERSION 5.00
Begin VB.Form frmTransactions 
   Caption         =   "Transaction Manager"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   9285
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "Close Transaction Manager"
      End
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = ShuttingDown
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PutWindowPosition Me, True
    fMain.mnuTransactions.Checked = False
End Sub


Private Sub mnuClose_Click()
    fMain.mnuTransactions.Checked = False
    Me.Hide
End Sub


