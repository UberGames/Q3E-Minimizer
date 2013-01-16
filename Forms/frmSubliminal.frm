VERSION 5.00
Begin VB.Form frmSubliminal 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Q3E_Minimizer_151"
   ClientHeight    =   2250
   ClientLeft      =   645
   ClientTop       =   600
   ClientWidth     =   3555
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picMOHAABOff 
      Height          =   315
      Left            =   0
      Picture         =   "frmSubliminal.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   42
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picMOHAASOff 
      Height          =   315
      Left            =   3240
      Picture         =   "frmSubliminal.frx":058A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   41
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picMOHAAOff 
      Height          =   315
      Left            =   2880
      Picture         =   "frmSubliminal.frx":0B14
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   40
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picMOHAABOn 
      Height          =   315
      Left            =   0
      Picture         =   "frmSubliminal.frx":109E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   39
      Top             =   720
      Width           =   315
   End
   Begin VB.PictureBox picMOHAASOn 
      Height          =   315
      Left            =   3240
      Picture         =   "frmSubliminal.frx":1628
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   38
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picMOHAAOn 
      Height          =   315
      Left            =   2880
      Picture         =   "frmSubliminal.frx":1BB2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   37
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picQ4Off 
      Height          =   315
      Left            =   2520
      Picture         =   "frmSubliminal.frx":213C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   36
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picGOiMOff 
      Height          =   315
      Left            =   2160
      Picture         =   "frmSubliminal.frx":26C6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   35
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picAliceOff 
      Height          =   315
      Left            =   1800
      Picture         =   "frmSubliminal.frx":2C50
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   34
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picQ4On 
      Height          =   315
      Left            =   2520
      Picture         =   "frmSubliminal.frx":31DA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   33
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picGOiMOn 
      Height          =   315
      Left            =   2160
      Picture         =   "frmSubliminal.frx":3764
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   32
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picAliceOn 
      Height          =   315
      Left            =   1800
      Picture         =   "frmSubliminal.frx":3CEE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   31
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picCoDUOOff 
      Height          =   315
      Left            =   1440
      Picture         =   "frmSubliminal.frx":4278
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   30
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picD3Off 
      Height          =   315
      Left            =   1080
      Picture         =   "frmSubliminal.frx":4802
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   29
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picCoDOff 
      Height          =   315
      Left            =   360
      Picture         =   "frmSubliminal.frx":4D8C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picSoF2MPOff 
      Height          =   315
      Left            =   0
      Picture         =   "frmSubliminal.frx":5316
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picFAKKOff 
      Height          =   315
      Left            =   720
      Picture         =   "frmSubliminal.frx":58A0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   26
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox picSoF2SPOff 
      Height          =   315
      Left            =   3240
      Picture         =   "frmSubliminal.frx":5E2A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picETOff 
      Height          =   315
      Left            =   2880
      Picture         =   "frmSubliminal.frx":63B4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picJKAMPOff 
      Height          =   315
      Left            =   2520
      Picture         =   "frmSubliminal.frx":693E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picJKASPOff 
      Height          =   315
      Left            =   2160
      Picture         =   "frmSubliminal.frx":6EC8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picRTCWOff 
      Height          =   315
      Left            =   1800
      Picture         =   "frmSubliminal.frx":7452
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   21
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picEF2Off 
      Height          =   315
      Left            =   1440
      Picture         =   "frmSubliminal.frx":79DC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   20
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picEFOff 
      Height          =   315
      Left            =   1080
      Picture         =   "frmSubliminal.frx":7F66
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picJK2SPOff 
      Height          =   315
      Left            =   720
      Picture         =   "frmSubliminal.frx":84F0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picJK2MPOff 
      Height          =   315
      Left            =   360
      Picture         =   "frmSubliminal.frx":8A7A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   1200
      Width           =   315
   End
   Begin VB.PictureBox picCoDUOOn 
      Height          =   315
      Left            =   1440
      Picture         =   "frmSubliminal.frx":9004
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picD3On 
      Height          =   315
      Left            =   1080
      Picture         =   "frmSubliminal.frx":958E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   360
      Width           =   315
      Begin VB.PictureBox Picture29 
         Height          =   255
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   15
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox picFAKKOn 
      Height          =   315
      Left            =   720
      Picture         =   "frmSubliminal.frx":9B18
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox PicCoDOn 
      Height          =   315
      Left            =   360
      Picture         =   "frmSubliminal.frx":A0A2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picSoF2MPOn 
      Height          =   315
      Left            =   0
      Picture         =   "frmSubliminal.frx":A62C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   360
      Width           =   315
   End
   Begin VB.PictureBox picSoF2SPOn 
      Height          =   315
      Left            =   3240
      Picture         =   "frmSubliminal.frx":ABB6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picETOn 
      Height          =   315
      Left            =   2880
      Picture         =   "frmSubliminal.frx":B140
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picJKAMPOn 
      Height          =   315
      Left            =   2520
      Picture         =   "frmSubliminal.frx":B6CA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picJKASPOn 
      Height          =   315
      Left            =   2160
      Picture         =   "frmSubliminal.frx":BC54
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picRTCWOn 
      Height          =   315
      Left            =   1800
      Picture         =   "frmSubliminal.frx":C1DE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   0
      Width           =   315
   End
   Begin VB.Timer tmrMain 
      Interval        =   10
      Left            =   3120
      Top             =   720
   End
   Begin VB.PictureBox picEF2On 
      Height          =   315
      Left            =   1440
      Picture         =   "frmSubliminal.frx":C768
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picEFOn 
      Height          =   315
      Left            =   1080
      Picture         =   "frmSubliminal.frx":CCF2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picJK2SPOn 
      Height          =   315
      Left            =   720
      Picture         =   "frmSubliminal.frx":D27C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picJK2MPOn 
      Height          =   315
      Left            =   360
      Picture         =   "frmSubliminal.frx":D806
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picQ3On 
      Height          =   315
      Left            =   0
      Picture         =   "frmSubliminal.frx":DD90
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picQ3Off 
      Height          =   315
      Left            =   0
      Picture         =   "frmSubliminal.frx":E31A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   1200
      Width           =   315
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopAbout 
         Caption         =   "&Main..."
      End
      Begin VB.Menu mPopRestore 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mPopResoRestore 
         Caption         =   "&Restore Resolution..."
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmSubliminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' This is the brunt of the Q3E Minimizer's execution,
' where the primary inner workings of the program are
' situated. These include initializing the settings,
' loading the program into the system tray, and the
' set-up to scan for active games.
'==================================================

Option Explicit
Dim formTrans As Boolean

Private Sub Form_Load()
    On Error Resume Next
    
    'If Q3E Minimizer is already open, don't open another copy. :P
    'having two copies open would screw up the hotkeys bigtime. >.<
    If App.PrevInstance = True Then
        MsgBox ("Q3E Minimizer is already open!"), vbExclamation
        End
    End If
    
    formTrans = False
    InitPublicVars

    Dim strRA() As String

    'the form must be fully visible before calling Shell_NotifyIcon

    'Thanks to Microsoft for their helpful page on how to get your
    'program to show up in the system tray. :)
    mPopResoRestore.Visible = False
    
    Me.Hide
    Me.Refresh
    
    'Set up system tray entry
    With nid
        .cbSize = Len(nid)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = frmSubliminal.picQ3Off.picture
        .szTip = "Q3E Minimizer: INACTIVE" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid

    'This code here saves and retrieves the settings in Q3E Minimizer in a registry store.
    If ReadSettings("HideQ3Keys") = "" Then
        Call WriteSettings("HideQ3Keys", "2 90")
    End If
    'Chop up registry entry and assign both values to the hotkey struct
    strRA = Split(ReadSettings("HideQ3Keys"), " ")
    hotKeysData(HK_MINIMIZE).ShiftFlags = Int(strRA(0))
    hotKeysData(HK_MINIMIZE).KeyASCII = Int(strRA(1))

    If ReadSettings("ShowQ3Keys") = "" Then
        Call WriteSettings("ShowQ3Keys", "3 90")
    End If
    'Chop up registry entry and assign both values to the hotkey struct
    strRA = Split(ReadSettings("ShowQ3Keys"), " ")
    hotKeysData(HK_RESTORE).ShiftFlags = Int(strRA(0))
    hotKeysData(HK_RESTORE).KeyASCII = Int(strRA(1))
    
    'Just in case some weird users want to use the same key to toggle the screen lol ;)
    If hotKeysData(HK_MINIMIZE).KeyASCII = hotKeysData(HK_RESTORE).KeyASCII And _
        hotKeysData(HK_MINIMIZE).ShiftFlags = hotKeysData(HK_RESTORE).ShiftFlags Then
        OnlyUsingOneHotKey = True
    End If
    
    'Their previously selected game
    If ReadSettings("CurrentGame") = "" Then
        Call WriteSettings("CurrentGame", AUTOMATIC_DETECT)
    End If
    frmMain.cmbGamesList.ListIndex = ReadSettings("CurrentGame")
    
    'The 'Hide this window on start up' checkbox value
    If ReadSettings("HideWinOnStart") = "" Then
        Call WriteSettings("HideWinOnStart", "0")
    End If
    frmMain.chkShowWinCheck.Value = ReadSettings("HideWinOnStart")

    'The 'Start on Windows Boot-up' Checkbox
    If ReadSettings("StartOnBoot") = "" Then
        Call WriteSettings("StartOnBoot", "0")
    End If
    frmMain.chkRebootCheck.Value = ReadSettings("StartOnBoot")

    'The DSM Checkbox
    If ReadSettings("DynaScreen") = "" Then
        Call WriteSettings("DynaScreen", "0")
    End If
    frmMain.chkDynaScreenResMod.Value = ReadSettings("DynaScreen")
    
    'The 'Minimize if game loses focus' Checkbox
    If ReadSettings("MinimizeOnUnFocus") = "" Then
        Call WriteSettings("MinimizeOnUnFocus", "0")
    End If
    frmMain.chkFocusMinimize.Value = ReadSettings("MinimizeOnUnFocus")
    
    'Sometimes the prog loads up the DSM chk set to 2 (ie locked)
    If frmMain.chkDynaScreenResMod.Value > 1 Then
        frmMain.chkDynaScreenResMod.Value = 1
    End If
    
    If frmMain.cmbGamesList.Locked = True Then
        frmMain.cmbGamesList.Locked = False
    End If

    ' Subclassing the form to get the Windows callback msgs.
    ' Used for hotkeys and external prog messages
    glWinRet = SetWindowLong(frmSubliminal.hWnd, GWL_WNDPROC, AddressOf CallbackMsgs)

    If GetDPIY() <> "96" And GetDPIX() <> "96" Then
        resizeText
    End If

    'This checks if we're supposed to display the intro
    'window on program startup or not
    If ReadSettings("HideWinOnStart") = 1 Then
        frmMain.Visible = False
    Else
        frmMain.Visible = True
    End If
    
    'Handle any command line arguments
    HandleCmdLine (Command$)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'this procedure receives the callbacks from the System Tray icon.
    Dim result As Long
    Dim Msg As Long

    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        Msg = x
    Else
        Msg = x / Screen.TwipsPerPixelX
    End If
    
    Select Case Msg
        Case WM_LBUTTONUP        '514 if game is on, maximize, else show frmSettings
            GetCurrentGame
            If gamecon <> 0 Then
                MaximizeGame
            Else
                frmSettings.WindowState = vbNormal
                result = SetForegroundWindow(frmSettings.hWnd)
                frmSettings.Show
            End If
        Case WM_LBUTTONDBLCLK    '515 show frmMain
            frmSettings.WindowState = vbNormal
            result = SetForegroundWindow(frmMain.hWnd)
            frmMain.Show
        Case WM_RBUTTONUP        '517 display popup menu
            result = SetForegroundWindow(Me.hWnd)
            Me.PopupMenu Me.mPopupSys
    End Select
End Sub

Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
       
    If retVal0 = True Then
        UnregisterHotKey Me.hWnd, HK_MINIMIZE
    End If
    
    ' If second hotkey is registered then unregister it.
    If retVal1 = True Then
        UnregisterHotKey Me.hWnd, HK_RESTORE
    End If

End Sub

Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    GetCurrentGame

    If gamecon <> 0 Then
        MaximizeGame
    
        'If AutoGame detection was on, reset it
        If frmMain.cmbGamesList.Enabled = False Then
            frmMain.cmbGamesList.ListIndex = AUTOMATIC_DETECT
        End If
    End If
    
    Unload Me
    Unload frmMain
    Unload frmSettings
End Sub

Private Sub mPopRestore_Click()
    'called when the user clicks the popup menu Restore command
    Dim result As Long
    frmSettings.WindowState = vbNormal
    result = SetForegroundWindow(frmSettings.hWnd)
    frmSettings.Show
End Sub
      
Private Sub mPopAbout_Click()
    'called when the user clicks the popup menu Restore command
    Dim result As Long
    frmMain.WindowState = vbNormal
    result = SetForegroundWindow(frmMain.hWnd)
    frmMain.Show
End Sub

Private Sub tmrMain_Timer()
    On Error Resume Next

    'This was the only way to make Garrett's Transparent checkbox code work.
    'Simply setting TransCheck to 1 beforehand did nothing. (That explains
    'the grey rectangle u see upon startup O_o )
    If formTrans = False And frmMain.Visible = True Then
        formTrans = True
        Call TransChkOpt(frmMain.chkDynaScreenResMod, frmMain, formTrans)
        Call TransChkOpt(frmMain.chkShowWinCheck, frmMain, formTrans)   '(chkTransCheck.Value = 1)
        Call TransChkOpt(frmMain.chkRebootCheck, frmMain, formTrans) '(chkTransCheck.Value = 1)
        Call TransChkOpt(frmMain.chkFocusMinimize, frmMain, formTrans)
        tmrMain.Interval = 1000
    End If

    GetCurrentGame 'updates the global variable gamecon
    DoIconz 'Change the system tray icon
    
    'If game is active
    If gamecon <> 0 Then
    
        'If haven't tried yet, register hotkeys
        'I tried using the retVal bools as a check b4, but if the system
        'had already used those keys, it put the program in an infinite loop O_o.
        If TriedToRegister = False Then 'Register the hotkeys if a game is running
            retVal0 = RegisterHotKeys(HK_MINIMIZE)
            
            If OnlyUsingOneHotKey = False Then
                retVal1 = RegisterHotKeys(HK_RESTORE)
            End If
        
            TriedToRegister = True
        End If
        
        'If the mimimize on unfocus checkbox, and we are unfocused
        'Bug fix: add a bool so it only does this check once :P
        'So we can do things like res change and no-flickering windows
        If frmMain.chkFocusMinimize.Value = 1 And GetForegroundWindow() <> gamecon Then
            MinimizeGame (True)
            'MsgBox ("Game: " & typDM_Game.dmPelsWidth & "x" & typDM_Game.dmPelsHeight & Chr(13) & "Screen: " & typDM_Screen.dmPelsWidth & "x" & typDM_Screen.dmPelsHeight)
        End If
        
        'Freeze the screen modifier.  If a user tried to
        'change it whilst a game was running, that would be bad
        'Also lock the cmbgames list. If the user switches game
        'with DSM active... it loses resolution data. >.<
        If frmMain.chkDynaScreenResMod.Value = 1 Then
            frmMain.cmbGamesList.Locked = True
            frmMain.chkDynaScreenResMod.Value = 2
            frmMain.chkDynaScreenResMod.Enabled = False
        End If
        
        'Slow down the timer. Now system performance counts!
        If frmMain.chkFocusMinimize.Value = 0 Then
            If tmrMain.Interval < 1000 Then
                tmrMain.Interval = 3000
            End If
        Else
            'if focusCheck box checked, make it really fast, so there is no delay :P
            'Probably a bit hacky
            tmrMain.Interval = 200
        End If

    Else 'No games are active
    
        'Set the timer refresh rate to quite fast
        'The player isn't ingame, so using up a tiny
        'bit more of system resources SHOULD be okay... :P
        If tmrMain.Interval <> 400 Then
            tmrMain.Interval = 400
        End If
        
        'Reset the register hotkeys check
        If TriedToRegister = True Then
            TriedToRegister = False
        End If
    
        'Unregister them pesky hotkeys so I can use ctrl-z in peace!
        'If hotkeys are registered, un-do it :P
        If retVal0 = True Then
            UnregisterHotKey frmSubliminal.hWnd, HK_MINIMIZE
            retVal0 = False
        End If
            
        If retVal1 = True Then
                UnregisterHotKey frmSubliminal.hWnd, HK_RESTORE
                retVal1 = False
        End If

        'If autogame detection did find a game, reset it now
        If frmMain.cmbGamesList.Enabled = False Then
            frmMain.cmbGamesList.ListIndex = AUTOMATIC_DETECT
            frmMain.cmbGamesList.Enabled = True
        End If

        'Perform auto detection check
        If frmMain.cmbGamesList.ListIndex = AUTOMATIC_DETECT Then
            'behold! The Automatic Game Detector!
            DoDetect 'Hehe pretty short actually... ;)
        End If
        
        'If screen resolution modifier was active, then
        'reset it
        If frmMain.chkDynaScreenResMod.Value > 1 Then
            frmMain.chkDynaScreenResMod.Value = 1
            frmMain.chkDynaScreenResMod.Enabled = True
            frmMain.cmbGamesList.Locked = False
        End If
        'No game detected.  Definately not minimized then :P
        If GameIsMinimized = True Then
            GameIsMinimized = False
        End If
    End If

End Sub

