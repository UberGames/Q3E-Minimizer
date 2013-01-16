Attribute VB_Name = "TiMFunctions"
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
' This chunk of code was made entirely by yours truly
' in order to make several aspects of the code more
' efficient.  It probably stands to be even more
' efficient, so it may be worth taking a look yourself. ;)
'
'==================================================

'Constant declaration for number of games Q3E has registered
'If you want to add another game, increase this by one
Public Const MAX_GAMES = 25
'Used for checks involving auto-detect
Public Const AUTOMATIC_DETECT = 0
'max number of cmd line parameters (and then their arguments)
'the program will accept
Public Const MAX_ARGUMENTS = 2

Public Enum HotKeyTypes
    HK_MINIMIZE = 0
    HK_RESTORE = 1
End Enum

'Structure where all the info needed for each game is kept. :)
Public Type gameParams
    gameFormal As String
    gameActual As String
    gameIconOn As Long
    gameIconOff As Long
End Type

'Structure which stores the information needed to save hotkey configs
Public Type hotKeyData
    ShiftFlags As Integer
    KeyASCII As Integer
End Type

'creates array for each game
Public games(MAX_GAMES) As gameParams
'create struct to hold hotkey data
Public hotKeysData(2) As hotKeyData

Public gamecon As Long
Public game As String
Public GameIsMinimized As Boolean

Public OnlyUsingOneHotKey As Boolean
Public TriedToRegister As Boolean
Public retVal0, retVal1 As Boolean

'************************************
'Public Sub MinimizeGame
'
'Minimizes a window that has the same
'title as the one currently selected.
'If DSM is active, also changes screen
'res back to registry settings
'************************************
Public Sub MinimizeGame(ByVal LostFocus As Boolean)
    If GameIsMinimized = False Then
            GetCurrentGame
    
            If gamecon <> 0 Then
                If frmMain.chkDynaScreenResMod.Value <> 0 Then
                    Call ChangeToDesktopResolution(LostFocus, False)
                End If
            
                'NB: ABSOLUTELY HAS to be called after the screen
                'change, or else we'll lose the data we had :P
                Call ShowWindow(gamecon, SW_HIDE) 'SW_HIDE
                Call SetForegroundWindow(GetDesktopWindow())
                GameIsMinimized = True
            End If
    End If
End Sub

Public Sub MaximizeGame()
        GetCurrentGame
    
        'as long as we aren't already focused
        'If GetForegroundWindow <> gamecon Then
    
            If gamecon <> 0 And GetForegroundWindow <> gamecon Then
                'Set Foreground win to Q3 first.  Else it screws up the
                'lose focus command
                'NB: Game HAS TO BE IN THE FOREGROUND WHEN MAXIMISING,
                'OR UNFOCUS WILL MINIMIZE IT AGAIN
                Call SetForegroundWindow(gamecon)
                DoEvents
                
                If frmMain.chkDynaScreenResMod.Value <> 0 Then
                    ChangeToGameResolution
                End If

                Call ShowWindow(gamecon, SW_RESTORE)
                GameIsMinimized = False
            End If
    'End If
    'MsgBox GameIsMinimized
End Sub

'************************************
'Public Sub GetCurrentGame
'
'Performs a search for any active windows
'that have a title matching the game
'currently selected.
'It then puts the result in the gamecon
'variable which can be accessed anywhere.
'************************************
Public Sub GetCurrentGame()
    game = games(frmMain.cmbGamesList.ListIndex).gameActual
    
    gamecon = FindWindow(vbNullString, game)
End Sub

'******************************************************
'Public Sub InitPublicVars
'
'The tedious task of manually entering
'in the data needed for each separate
'game. You can add new games here. :)
'The order can be shuffled, however it
'would be wise to leave the automatic
'one at the top.
'Format:
'gameFormal = name of game as appers in drop down menu
'gameActual = name of game window, which is minimized
'gameIconOn = the icon file that is used when game is active
'gameIconOff = the icon file that is used when game is inactive
'*****************************************************

Public Sub InitPublicVars()

        With games(0)
            .gameFormal = "-=Automatic Game Detection=-"
            .gameActual = "Quake 3: Arena"
            .gameIconOn = frmSubliminal.picQ3On.picture
            .gameIconOff = frmSubliminal.picQ3Off.picture
        End With
        
        With games(1)
            .gameFormal = "Quake III Arena"
            .gameActual = "Quake 3: Arena"
            .gameIconOn = frmSubliminal.picQ3On.picture
            .gameIconOff = frmSubliminal.picQ3Off.picture
        End With
        
        With games(2)
            .gameFormal = "Quake 4"
            .gameActual = "Quake4"
            .gameIconOn = frmSubliminal.picQ4On.picture
            .gameIconOff = frmSubliminal.picQ4Off.picture
        End With
        
        With games(3)
            .gameFormal = "Doom 3"
            .gameActual = "DOOM 3"
            .gameIconOn = frmSubliminal.picD3On.picture
            .gameIconOff = frmSubliminal.picD3Off.picture
        End With
        
        With games(4)
            .gameFormal = "Heavy Metal: F.A.K.K²"
            .gameActual = "Heavy Metal: FAKK2"
            .gameIconOn = frmSubliminal.picFAKKOn.picture
            .gameIconOff = frmSubliminal.picFAKKOff.picture
        End With
  
        With games(5)
            .gameFormal = "American McGee's Alice"
            .gameActual = "American McGee's Alice"
            .gameIconOn = frmSubliminal.picAliceOn.picture
            .gameIconOff = frmSubliminal.picAliceOff.picture
        End With

        With games(6)
            .gameFormal = "Star Trek Voyager: Elite Force (SP)"
            .gameActual = "Star Trek: Voyager"
            .gameIconOn = frmSubliminal.picEFOn.picture
            .gameIconOff = frmSubliminal.picEFOff.picture
        End With

        With games(7)
            .gameFormal = "Star Trek Voyager: Elite Force (MP)"
            .gameActual = "Star Trek: Voyager - Holomatch"
            .gameIconOn = frmSubliminal.picEFOn.picture
            .gameIconOff = frmSubliminal.picEFOff.picture
        End With
        
        With games(8)
            .gameFormal = "Star Trek: Elite Force II"
            .gameActual = "Elite Force II"
            .gameIconOn = frmSubliminal.picEF2On.picture
            .gameIconOff = frmSubliminal.picEF2Off.picture
        End With
        
        With games(9)
            .gameFormal = "Call of Duty (SP)"
            .gameActual = "Call of Duty"
            .gameIconOn = frmSubliminal.PicCoDOn.picture
            .gameIconOff = frmSubliminal.picCoDOff.picture
        End With
        
        With games(10)
            .gameFormal = "Call of Duty (MP)"
            .gameActual = "Call of Duty Multiplayer"
            .gameIconOn = frmSubliminal.PicCoDOn.picture
            .gameIconOff = frmSubliminal.picCoDOff.picture
        End With
        
        With games(11)
            .gameFormal = "CoD: United Offensive (SP)"
            .gameActual = "CoD:United Offensive"
            .gameIconOn = frmSubliminal.picCoDUOOn.picture
            .gameIconOff = frmSubliminal.picCoDUOOff.picture
        End With
        
        With games(12)
            .gameFormal = "CoD: United Offensive (MP)"
            .gameActual = "CoD:United Offensive Multiplayer"
            .gameIconOn = frmSubliminal.picCoDUOOn.picture
            .gameIconOff = frmSubliminal.picCoDUOOff.picture
        End With
        
        With games(13)
            .gameFormal = "Return to Castle Wolfenstein (SP + MP)"
            .gameActual = "Wolfenstein"
            .gameIconOn = frmSubliminal.picRTCWOn.picture
            .gameIconOff = frmSubliminal.picRTCWOff.picture
        End With
        
        With games(14)
            .gameFormal = "RTCW: Enemy Territory"
            .gameActual = "Enemy Territory"
            .gameIconOn = frmSubliminal.picETOn.picture
            .gameIconOff = frmSubliminal.picETOff.picture
        End With
        
        With games(15)
            .gameFormal = "Game Over in Machinimation"
            .gameActual = "Game Over"
            .gameIconOn = frmSubliminal.picGOiMOn.picture
            .gameIconOff = frmSubliminal.picGOiMOff.picture
        End With
        
        With games(16)
            .gameFormal = "Soldier of Fortune II: Double Helix (SP)"
            .gameActual = "Soldier of Fortune 2 : Double Helix"
            .gameIconOn = frmSubliminal.picSoF2SPOn.picture
            .gameIconOff = frmSubliminal.picSoF2SPOff.picture
        End With
        
        With games(17)
            .gameFormal = "Soldier of Fortune II: Double Helix (MP)"
            .gameActual = "SoF2 MP"
            .gameIconOn = frmSubliminal.picSoF2MPOn.picture
            .gameIconOff = frmSubliminal.picSoF2MPOff.picture
        End With
        
        With games(18)
            .gameFormal = "Medal of Honor Allied Assault"
            .gameActual = "Medal of Honor Allied Assault"
            .gameIconOn = frmSubliminal.picMOHAAOn.picture
            .gameIconOff = frmSubliminal.picMOHAAOff.picture
        End With
        
        With games(19)
            .gameFormal = "Medal of Honor Allied Assault Spearhead"
            .gameActual = "Medal of Honor Allied Assault(TM) Spearhead"
            .gameIconOn = frmSubliminal.picMOHAASOn.picture
            .gameIconOff = frmSubliminal.picMOHAASOff.picture
        End With
        
        With games(20)
            .gameFormal = "Medal of Honor Allied Assault Breakthrough"
            .gameActual = "Medal of Honor Allied Assault(TM) Breakthrough"
            .gameIconOn = frmSubliminal.picMOHAABOn.picture
            .gameIconOff = frmSubliminal.picMOHAABOff.picture
        End With
        
        With games(21)
            .gameFormal = "Star Wars: Jedi Knight II: Jedi Outcast (SP)"
            .gameActual = "Jedi Knight 2"
            .gameIconOn = frmSubliminal.picJK2SPOn.picture
            .gameIconOff = frmSubliminal.picJK2SPOff.picture
        End With
        
        With games(22)
            .gameFormal = "Star Wars: Jedi Knight II: Jedi Outcast (MP)"
            .gameActual = "Jedi Knight 2: Jedi Outcast MP"
            .gameIconOn = frmSubliminal.picJK2MPOn.picture
            .gameIconOff = frmSubliminal.picJK2MPOff.picture
        End With
   
        With games(23)
            .gameFormal = "Star Wars: Jedi Knight: Jedi Academy (SP)"
            .gameActual = "Jedi Knight®: Jedi Academy"
            .gameIconOn = frmSubliminal.picJKASPOn.picture
            .gameIconOff = frmSubliminal.picJKASPOff.picture
        End With
        
        With games(24)
            .gameFormal = "Star Wars: Jedi Knight: Jedi Academy (MP)"
            .gameActual = "Jedi Knight®: Jedi Academy (MP)"
            .gameIconOn = frmSubliminal.picJKAMPOn.picture
            .gameIconOff = frmSubliminal.picJKAMPOff.picture
        End With
End Sub
'******************************************************
'Public Function DoDetect
'
'Called every few seconds in the timer control.
'This scans through the list of games, checking for
'if each one is active or not.  If it finds an active
'game, it switches the settings to that game
'*****************************************************

Public Function DoDetect()
        
        For i = 0 To (MAX_GAMES - 1) Step 1
            gamecon = FindWindow(vbNullString, games(i).gameActual)
            
            If gamecon <> 0 And frmMain.cmbGamesList.Enabled = True Then
                frmMain.cmbGamesList.Text = games(i).gameFormal
                frmMain.cmbGamesList.Enabled = False
                Exit For
            Else
                Shell_NotifyIcon NIM_MODIFY, nid
            
                With nid
                    .hIcon = frmSubliminal.picQ3Off.picture
                    .szTip = "Q3E Minimizer - Auto Game Detection Enabled" & vbNullChar
                End With
            End If
        Next
End Function

'******************************************************
'Public Function DoIconz
'
'Updates the icon in the system tray to match the
'currently selected game
'******************************************************

Public Function DoIconz()
    Shell_NotifyIcon NIM_MODIFY, nid

    If nid.hIcon <> games(frmMain.cmbGamesList.ListIndex).gameIconOn Or nid.hIcon <> _
        games(frmMain.cmbGamesList.ListIndex).gameIconOff Then

        GetCurrentGame

        With nid
            If gamecon <> 0 Then
                .hIcon = games(frmMain.cmbGamesList.ListIndex).gameIconOn
                If frmMain.cmbGamesList.ListIndex <> AUTOMATIC_DETECT Then
                    .szTip = frmMain.cmbGamesList.Text & " - ACTIVE" & vbNullChar
                End If
            Else
                .hIcon = games(frmMain.cmbGamesList.ListIndex).gameIconOff
                If frmMain.cmbGamesList.ListIndex <> AUTOMATIC_DETECT Then
                    .szTip = frmMain.cmbGamesList.Text & " - INACTIVE" & vbNullChar
                End If
            End If
       End With
    End If
End Function

'******************************************************
'Public Function RegisterHotKeys
'
'When called, registers the hotkey of a specific index
'******************************************************

Public Function RegisterHotKeys(hotKeyIndex As Integer) As Boolean
    Dim result As Boolean
    Dim ShiftFlags As Integer
    Dim endFlags As Integer
    
    'Pfft. Just in case... >.<
    If hotKeysData(hotKeyIndex).ShiftFlags = 0 And _
        hotKeysData(hotKeyIndex).KeyASCII = 0 Then
        RegisterHotKeys = False
        Exit Function
    End If
    
    ShiftFlags = hotKeysData(hotKeyIndex).ShiftFlags
    
    'Weird complicated conditional thingy needed in order to convert
    'the Shift flags to the hot key shift flag values :P
    If (ShiftFlags And vbCtrlMask) Then
        endFlags = MOD_CTRL
        If (ShiftFlags And vbShiftMask) Then
            endFlags = MOD_SHIFTCONTROL
            If (ShiftFlags And vbAltMask) Then
                endFlags = MOD_CONTROLALTSHIFT
                GoTo Finish
            End If
            GoTo Finish
        End If
        If (ShiftFlags And vbAltMask) Then
            endFlags = MOD_CONTROLALT
            GoTo Finish
        End If
    End If
    
    If (ShiftFlags And vbAltMask) Then
        endFlags = MOD_ALT
        If (ShiftFlags And vbShiftMask) Then
            endFlags = MOD_ALTSHIFT
            GoTo Finish
        End If
        GoTo Finish
    End If
    
    If (ShiftFlags And vbShiftMask) Then
        endFlags = MOD_SHFT
    End If
           
Finish:
    result = RegisterHotKey(frmSubliminal.hWnd, hotKeyIndex, endFlags, hotKeysData(hotKeyIndex).KeyASCII)

    If result = False Then
        Dim cmdResult As String
        Select Case hotKeyIndex
            Case HK_MINIMIZE
                cmdResult = "hide"
            Case HK_RESTORE
                cmdResult = "show"
        End Select
    
        MsgBox "Could not register key sequence to " & cmdResult & " game. Another program may already be using that key combination. Please enter in a different key combination for it.", vbCritical
    End If
    
    RegisterHotKeys = result
    
End Function

'******************************************************
'Public Function HandleCmdLine
'
'Used to handle any command line parameters the user
'may have started the program up with.
'NB ASCII number 34 = "
'******************************************************

Public Function HandleCmdLine(Command As String)

    On Error Resume Next
    Dim strRA(MAX_ARGUMENTS) As String
    Dim i As Integer
    Dim fileName As String
    Dim shellRes As Long
    Dim strStart, strLength As Integer
    Dim tempString As String
    strStart = 1
    i = 0
    
    'Divide the Command text into an array where each word between
    'spaces is a separate cell
    'strRA = Split(Command, " ")
   
    'parse the command string into multiple cell arrays.
    'Why I don't use the Split command is because if there
    'are spaces in the file route, it chops that up too. Bad! >.<
    Do While 1
    
        'if there's a space at the beginning, screw it
        If Mid(Command, strStart, 1) = " " Then
            strStart = strStart + 1
        End If
        
        If Mid(Command, strStart, 1) = Chr(34) Then
            strLength = InStr(strStart + 1, Command, Chr(34)) - strStart + 1
        Else
            strLength = InStr(strStart, Command, " ") - strStart
        End If
        
        tempString = Mid(Command, strStart, strLength)
        strRA(i) = tempString
        
        strStart = strStart + strLength
        
        i = i + 1
        
        If strLength < 0 Or i > UBound(strRA) Then
           Exit Do
        End If
    Loop
    
    'Go through and check each cell for command matches
    For i = LBound(strRA) To UBound(strRA)
        Select Case LCase(strRA(i))
           'user wants to start another program
            Case "-run"
               'if this was the last cell... that's bad... no file route
                If i = UBound(strRA) Then
                    MsgBox ("No file specified to start up")
                Else
                    i = i + 1
                End If
               'if invalid chars start off the file route
                If Left(strRA(i), 1) = "-" Or Left(strRA(i), 1) = "/" Then
                    MsgBox "Invalid file name specified"
                Else
                    On Error GoTo fileError
                    fileName = strRA(i)
                    frmMain.Hide
                    frmSettings.Hide
                   shellRes = Shell(fileName, vbNormalFocus)
                End If
        End Select
    Next i
    Exit Function
fileError:
    MsgBox ("Warning: A problem was detected when trying to open the file specified in the command line.  It may be incompatible or may not exist.")
    frmMain.Show
End Function

'******************************************************
'Public Function ASCIIChars
'
'Used to assign a name to any of the secondary keys
'that may have been pressed.
'******************************************************
Public Function ASCIIChars(ascii As Integer) As String
    Dim AscString As String

    If ascii > 0 Then
        Select Case ascii
            Case 12:    AscString = "NumPad ="
            Case 32:    AscString = "Space"
            Case 37:    AscString = "Left"
            Case 38:    AscString = "Up"
            Case 39:    AscString = "Right"
            Case 40:    AscString = "Down"
            Case 144:   AscString = "NumLock"
            Case 111:   AscString = "NumPad /"
            Case 106:   AscString = "NumPad *"
            Case 109:   AscString = "NumPad -"
            Case 97:    AscString = "NumPad 1"
            Case 98:    AscString = "NumPad 2"
            Case 99:    AscString = "NumPad 3"
            Case 100:   AscString = "NumPad 4"
            Case 101:   AscString = "Numpad 5"
            Case 102:   AscString = "NumPad 6"
            Case 103:   AscString = "NumPad 7"
            Case 104:   AscString = "NumPad 8"
            Case 105:   AscString = "NumPad 9"
            Case 107:   AscString = "NumPad +"
            Case 96:    AscString = "NumPad 0"
            Case 110:   AscString = "NumPad ."
            Case 112:   AscString = "F1"
            Case 113:   AscString = "F2"
            Case 114:   AscString = "F3"
            Case 115:   AscString = "F4"
            Case 116:   AscString = "F5"
            Case 117:   AscString = "F6"
            Case 118:   AscString = "F7"
            Case 119:   AscString = "F8"
            Case 120:   AscString = "F9"
            Case 121:   AscString = "F10"
            Case 122:   AscString = "F11"
            Case 123:   AscString = "F12"
            Case 124:   AscString = "F13"
            Case 125:   AscString = "F14"
            Case 126:   AscString = "F15"
            Case 127:   AscString = "F16"
            Case 128:   AscString = "F17"
            Case 129:   AscString = "F18"
            Case 130:   AscString = "F19"
            Case 131:   AscString = "F20"
            Case 132:   AscString = "F21"
            Case 133:   AscString = "F22"
            Case 134:   AscString = "F23"
            Case 135:   AscString = "F24"
            Case 192:   AscString = "`"
            Case 189:   AscString = "-"
            Case 187:   AscString = "="
            Case 220:   AscString = "\"
            Case 8:     AscString = "Backspace"
            Case 45:    AscString = "Insert"
            Case 36:    AscString = "Home"
            Case 33:    AscString = "PageUp"
            Case 46:    AscString = "Delete"
            Case 35:    AscString = "End"
            Case 34:    AscString = "PageDown"
            Case 9:     AscString = "Tab"
            Case 219:   AscString = "["
            Case 221:   AscString = "]"
            Case 20:    AscString = "CapsLock"
            Case 186:   AscString = ";"
            Case 222:   AscString = "'"
            Case 188:   AscString = ","
            Case 190:   AscString = "."
            Case 191:   AscString = "/"
        End Select
            
        ASCIIChars = AscString
    End If
            
End Function

'******************************************************
'Public Function resizeText
'
'This is called if the screen's DPI is different.
'In order to compnsate for this, each control has to be run through
'a calculating function. O_o

'******************************************************

Public Function resizeText()
        frmMain.txtAbout.Visible = False
        frmMain.txtAboutScroll.Visible = True
        frmMain.txtAboutScroll.Text = frmMain.txtAbout.Text
        
        With frmMain
            .Height = EstimateFontSize(frmMain.Height)
            .Width = EstimateFontSize(frmMain.Width)
        End With
        
        With frmMain.chkDynaScreenResMod
            .FontSize = Round(EstimateFontSize(frmMain.chkDynaScreenResMod.FontSize))
            .Width = EstimateFontSize(frmMain.chkDynaScreenResMod.Width)
            .Height = EstimateFontSize(frmMain.chkDynaScreenResMod.Height)
            .Left = EstimateFontSize(frmMain.chkDynaScreenResMod.Left)
            .Top = EstimateFontSize(frmMain.chkDynaScreenResMod.Top)
        End With

        With frmMain.chkRebootCheck
            .FontSize = Round(EstimateFontSize(frmMain.chkRebootCheck.FontSize))
            .Width = EstimateFontSize(frmMain.chkRebootCheck.Width)
            .Height = EstimateFontSize(frmMain.chkRebootCheck.Height)
            .Left = EstimateFontSize(frmMain.chkRebootCheck.Left)
            .Top = EstimateFontSize(frmMain.chkRebootCheck.Top)
        End With
        
        With frmMain.chkShowWinCheck
            .FontSize = Round(EstimateFontSize(frmMain.chkShowWinCheck.FontSize))
            .Width = EstimateFontSize(frmMain.chkShowWinCheck.Width)
            .Height = EstimateFontSize(frmMain.chkShowWinCheck.Height)
            .Left = EstimateFontSize(frmMain.chkShowWinCheck.Left)
            .Top = EstimateFontSize(frmMain.chkShowWinCheck.Top)
        End With
        
        With frmMain.cmbGamesList
            .FontSize = Round(EstimateFontSize(frmMain.chkShowWinCheck.FontSize))
            .Width = EstimateFontSize(frmMain.chkShowWinCheck.Width)
            '.Height = EstimateFontSize(chkShowWinCheck.Height)
            .Left = EstimateFontSize(frmMain.cmbGamesList.Left)
            .Top = EstimateFontSize(frmMain.cmbGamesList.Top)
        End With
        
        With frmMain.btnOkay
            .FontSize = Round(EstimateFontSize(frmMain.btnOkay.FontSize))
            .Width = EstimateFontSize(frmMain.btnOkay.Width)
            .Height = EstimateFontSize(frmMain.btnOkay.Height)
            .Left = EstimateFontSize(frmMain.btnOkay.Left)
            .Top = EstimateFontSize(frmMain.btnOkay.Top)
        End With
        
        With frmMain.txtAboutScroll
            .FontSize = Round(EstimateFontSize(frmMain.txtAboutScroll.FontSize))
            .Width = EstimateFontSize(frmMain.txtAboutScroll.Width)
            .Height = EstimateFontSize(frmMain.txtAboutScroll.Height)
            .Left = EstimateFontSize(frmMain.txtAboutScroll.Left)
            .Top = EstimateFontSize(frmMain.txtAboutScroll.Top)
            '.ScrollBars = 2
        End With

        frmSettings.txtInstructionsScroll.Visible = True
        frmSettings.txtInstructions.Visible = False
    
        With frmSettings
            .Height = EstimateFontSize(frmSettings.Height)
            .Width = EstimateFontSize(frmSettings.Width)
        End With
        
        With frmSettings.btnOkay
            .FontSize = Round(EstimateFontSize(frmSettings.btnOkay.FontSize))
            .Width = EstimateFontSize(frmSettings.btnOkay.Width)
            .Height = EstimateFontSize(frmSettings.btnOkay.Height)
            .Left = EstimateFontSize(frmSettings.btnOkay.Left)
            .Top = EstimateFontSize(frmSettings.btnOkay.Top)
        End With

        With frmSettings.lblHide
            .FontSize = Round(EstimateFontSize(frmSettings.lblHide.FontSize))
            .Width = EstimateFontSize(frmSettings.lblHide.Width) + EstimateFontSize(frmSettings.lblHide.Width) * 1.03
            .Height = EstimateFontSize(frmSettings.lblHide.Height)
            .Left = EstimateFontSize(frmSettings.lblHide.Left) - (EstimateFontSize(frmSettings.lblHide.Left)) * 0.035
            .Top = EstimateFontSize(frmSettings.lblHide.Top)
        End With
        
        With frmSettings.lblRstore
            .FontSize = Round(EstimateFontSize(frmSettings.lblRstore.FontSize))
            .Width = EstimateFontSize(frmSettings.lblRstore.Width) + EstimateFontSize(frmSettings.lblRstore.Width) * 1.03
            .Height = EstimateFontSize(frmSettings.lblRstore.Height)
            .Left = EstimateFontSize(frmSettings.lblRstore.Left) - (EstimateFontSize(frmSettings.lblRstore.Left)) * 0.035
            .Top = EstimateFontSize(frmSettings.lblRstore.Top)
        End With
        
        With frmSettings.txtHideKeyz
            .FontSize = Round(EstimateFontSize(frmSettings.txtHideKeyz.FontSize))
            .Width = EstimateFontSize(frmSettings.txtHideKeyz.Width)
            .Height = EstimateFontSize(frmSettings.txtHideKeyz.Height)
            .Left = EstimateFontSize(frmSettings.txtHideKeyz.Left)
            .Top = EstimateFontSize(frmSettings.txtHideKeyz.Top)
        End With
        
        With frmSettings.txtInstructionsScroll
            .FontSize = Round(EstimateFontSize(frmSettings.txtInstructions.FontSize))
            .Width = EstimateFontSize(frmSettings.txtInstructions.Width)
            .Height = EstimateFontSize(frmSettings.txtInstructions.Height)
            .Left = EstimateFontSize(frmSettings.txtInstructions.Left)
            .Top = EstimateFontSize(frmSettings.txtInstructions.Top)
        End With

        With frmSettings.txtRestoreKeyz
            .FontSize = Round(EstimateFontSize(frmSettings.txtRestoreKeyz.FontSize))
            .Width = EstimateFontSize(frmSettings.txtRestoreKeyz.Width)
            .Height = EstimateFontSize(frmSettings.txtRestoreKeyz.Height)
            .Left = EstimateFontSize(frmSettings.txtRestoreKeyz.Left)
            .Top = EstimateFontSize(frmSettings.txtRestoreKeyz.Top)
        End With

End Function
