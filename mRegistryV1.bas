Attribute VB_Name = "mRegistry"
'// Provides registry code.
'// Copyright (c)2002 Richard Holyoak.
'// Contact rholyoak@bigfoot.com
'// Requires RegistrationDatabase V1

Option Explicit

Public RegistryError As Integer

Private Const mstrVendor As String = "Software\Richard Holyoak\"

'// Registry keys
Public Const REG_OPTIONS As String = "Options"
Public Const REG_SETTINGS As String = "Settings"

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259
Public Sub RegPut(Section As String, Key As String, vSetting As Variant, lValueType As Long)
    On Error Resume Next
    'If App.ProductName = "" Then MsgBox "Error: No Product Name for this project!", vbCritical, "GetReg()": End
    
    Dim hKey As Long
    
    RegistryError = ERROR_NONE
    
    Select Case Section
        Case ""
            hKey = HKEY_CURRENT_USER
        Case REG_OPTIONS
            hKey = HKEY_CURRENT_USER
        Case REG_SETTINGS
            hKey = HKEY_CURRENT_USER
        Case Else
            hKey = HKEY_CURRENT_USER
    End Select
    
    CreateNewKey hKey, mstrVendor & App.Title & "\" & Section
    RegistryError = SetKeyValue(hKey, mstrVendor & App.Title & "\" & Section, Key, vSetting, lValueType)
    
End Sub

Public Function RegGet(Section As String, Key As String, Optional Default As String) As Variant
    Dim lRetValue As Long, hKey As Long
    On Error Resume Next
    
    RegistryError = ERROR_NONE
    Select Case Section
        Case ""
            hKey = HKEY_CURRENT_USER
        Case REG_OPTIONS
            hKey = HKEY_CURRENT_USER
        Case REG_SETTINGS
            hKey = HKEY_CURRENT_USER
        Case Else
            hKey = HKEY_CURRENT_USER
    End Select
    
    'If App.ProductName = "" Then MsgBox "Error: No Product Name for this project!", vbCritical, "GetReg()": End
    RegGet = QueryValue(hKey, mstrVendor & App.Title & "\" & Section, Key, lRetValue)
    If lRetValue <> ERROR_NONE Then RegGet = Default
End Function

Public Sub RegDel(Optional Section As Variant)
    Dim hKey As Long, lRet As Long
    
    RegistryError = ERROR_NONE
    
    If IsMissing(Section) Then
        Section = mstrVendor & App.Title
    Else
        Section = mstrVendor & App.Title & "\" & Section
    End If
    
    hKey = HKEY_CURRENT_USER
DeleteSetting Section
'    lRet = RegDeleteKey(hKey, Section)

End Sub
Public Function DeleteSetting(ByVal Section As String, Optional ByVal Key As String = "") As Boolean
   ' Section   Required. String expression containing the name of the section where the key setting
   '           is being deleted. If only section is provided, the specified section is deleted along
   '           with all related key settings.
   ' Key       Optional. String expression containing the name of the key setting being deleted.
   Dim nRet As Long
   Dim hKey As Long

   If Len(Key) Then
      ' Open key
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, SubKey(Section), 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Set appropriate value for default query
         If Key = "*" Then Key = vbNullString
         ' Delete the requested value
         nRet = RegDeleteValue(hKey, Key)
         Call RegCloseKey(hKey)
      End If
   Else
      ' Open parent key
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, SubKey(), 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Attempt to delete whole section
         nRet = RegDeleteKey(hKey, Section)
         Call RegCloseKey(hKey)
      End If
   End If
   DeleteSetting = (nRet = ERROR_SUCCESS)
End Function

' ********************************************
'  Private Methods
' ********************************************
Private Function SubKey(Optional ByVal Section As String = "") As String
   ' Build SubKey from known values
   SubKey = mstrVendor '& App.Title
   If Len(Section) Then
      SubKey = SubKey & "\" & Section
   End If
End Function


