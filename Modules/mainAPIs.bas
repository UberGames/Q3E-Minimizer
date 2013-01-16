Attribute VB_Name = "mainAPIs"
'=======================================================================
'Copyright (C) 2006 Timothy 'TiM' Oliver
'
'This file is part of the Q3E Minimizer v1.51 source code.
'
'Q3E Minimizer is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'Q3E Minimizer is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with Q3E Minimizer; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
'=======================================================================

'==================================================
' Q3E Minimizer 1.51 by UberGames
' Developed by Timothy 'TiM' Oliver
'
' The following code lists the main Windows API
' functions used in Q3E minimizer.  The first set
' declares the function used to actually minimize/
' maximize the game window, and the second is used
' to create the instance of the program in the Windows
' system tray.
' The third set read and write data to a text file
' for the purpose of saving the program settings.
'
'==================================================

'--------------------------------------------------

'****************************
'This probably takes up some space, but I've left a lot of unused
'functions in, just in case someone might find them useful. ;)
'How much extra can a few lines take...? :)
'****************************

' Windows related API calls---------------------------
      
Public Const SW_HIDE = 0
'Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
'Public Const SW_NORMAL = 1
'Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9

'------------------------------------------------
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function IsIconic Lib "user32" (ByRef hWnd As Long) As Boolean
Public Declare Function IsZoomed Lib "user32" (ByRef hWnd As Long) As Boolean
'Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
'Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
    
'------------------------------------------
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
       
' System Tray related API Calls---------------------------
      
'user defined type required by Shell_NotifyIcon API call
     
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid As NOTIFYICONDATA

'-------------------------------------------------------

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadSettings(KeyName As String) As String
    'Dim sRet As String
    'sRet = String(255, Chr(0))
    'ReadSettings = Left(sRet, GetPrivateProfileString("Settings", ByVal KeyName, "", sRet, Len(sRet), App.Path & "\settings.dat"))
    ReadSettings = GetSetting("Q3E Minimizer v1.50", "Program", KeyName)
    
End Function

Function WriteSettings(sKeyName As String, sNewString As String) As Integer
    'Dim r
    'r = WritePrivateProfileString("Settings", sKeyName, sNewString, App.Path & "\settings.dat")
    SaveSetting "Q3E Minimizer v1.50", "Program", sKeyName, sNewString

End Function
