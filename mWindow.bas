Attribute VB_Name = "mWindow"
'// Provides portable modular code for window operations.
'// Copyright (c)2002 Richard Holyoak.
'// Contact rholyoak@bigfoot.com

'// Requires at least modRegistryV1

Option Explicit

'// Registry keys
Public Const REG_WINDOW As String = "Windows"

Public Sub GetWindowPosition(frmForm As Form)
    Dim h As Single, w As Single, t As Single, l As Single
    On Error Resume Next
    
    With frmForm
    
        '// Get the initial position, width and height of the window
        w = RegGet(REG_WINDOW, .Caption & " Width", .Width)
        h = RegGet(REG_WINDOW, .Caption & " Height", .Height)
        l = RegGet(REG_WINDOW, .Caption & " Left", (fMain.Width - w) / 2)
        t = RegGet(REG_WINDOW, .Caption & " Top", (fMain.Height - h) / 2)
        
        '// Make sure the origin of the window is within the screen boundaries
        If Screen.Width < l Then l = (fMain.Width - w) / 2
        If Screen.Height < t Then t = (fMain.Height - h) / 2
        If l < 0 Then l = 0
        If t < 0 Then t = 0
        
        '// Setup the window
        .Move l, t, w, h
        .WindowState = RegGet(REG_WINDOW, .Caption & " WindowState", .WindowState)
        
    End With
    
End Sub

Public Sub PutWindowPosition(frmForm As Form, Optional Metrics As Boolean)
    On Error Resume Next
    With frmForm
        RegPut REG_WINDOW, .Caption & " Left", .Left, REG_DWORD
        RegPut REG_WINDOW, .Caption & " Top", .Top, REG_DWORD
        If Metrics Then RegPut REG_WINDOW, .Caption & " Width", .Width, REG_DWORD
        If Metrics Then RegPut REG_WINDOW, .Caption & " Height", .Height, REG_DWORD
        RegPut REG_WINDOW, .Caption & " WindowState", .WindowState, REG_DWORD
    End With
End Sub


