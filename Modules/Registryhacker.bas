Attribute VB_Name = "Registryhacker"

'==================================================
' Q3E Minimizer 1.51 by UberGames
' Developed by Timothy 'TiM' Oliver
'
' The following code lists the functions required
' by the code needed to modify the system registry
' so that Windows can be set to make it so Q3E
' Minimizer can be made to open up on Windows boot-up.
'
' Credit to manavo11 for the majority of this code.
'
'==================================================

'--------------------------------------------------

Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const ERROR_SUCCESS = 0&

Public Enum pvpHK
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Const pvpRunHKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"

Private Sub savestring(ByVal Hkey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If r = 87 Then
        DeleteValue Hkey, strPath, strValue
    End If
    r = RegCloseKey(keyhand)
End Sub

Private Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Public Function RunAtStartup(sAppTitle As String, strsAppName As String)
         savestring pvpHK.HKEY_CURRENT_USER, pvpRunHKey, sAppTitle, strsAppName
End Function

Public Function RemoveFromStartup(sAppTitle As String, strsAppName As String)
         DeleteValue pvpHK.HKEY_CURRENT_USER, pvpRunHKey, sAppTitle
End Function
