VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private cIni As Ini

'// Keys

'// Defaults

Private Sub Class_Initialize()
    Set cIni = New Ini
    cIni.IniFile = App.EXEName
    cIni.Vendor = App.Title
    
    With Me
    End With

End Sub

Private Sub Class_Terminate()
    Set cIni = Nothing
End Sub

Public Sub ViewIni()
    On Error GoTo Error_
    
    Shell "Notepad.exe """ & App.EXEName & ".ini" & """"
    
Exit_:
    Exit Sub

Error_:
    Alert Err, vbExclamation, "Options.ViewIni"
    Resume Exit_
    
End Sub
