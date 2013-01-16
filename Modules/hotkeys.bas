Attribute VB_Name = "Hotkeys"
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
' The following code is used to define the functions
' used to register and then subsequently execute the
' hotkey commands used in the program.
'
' This code was developed by Imran Zaheer and was
' made free for online download.
'
'==================================================

'--------------------------------------------------

Option Explicit

' ***********************************
' Author : Imran Zaheer
' Email  : imraanz@mail.com
' Web    : www.imraanz.com
' Y2K
' Module : Contains declarations and functions for
'          vHotKeys.

Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_HOTKEY = &H312
'Q3E Minimizer Subclass callback defines
Public Const WM_MINCALLBACK = &H313

Public Const ID_TOGGLE = 0 'Toggles between minimize and maximize
Public Const ID_MINIMIZE = 1 'Minimizes game
Public Const ID_MAXIMIZE = 2 'Maximizes game
Public Const ID_SAVEGAMERES = 3 'Saves the current resolution settings as the game's resolution
Public Const ID_GOTOSCREENRES = 4 'Goes to Desktop Screen Res
Public Const ID_GOTOGAMERES = 5 'Goes to InGame Res
'NB: There is no SaveDesktopRes since this set of data is stored directly in the registry
'and is accessed from there.

Public Const GWL_WNDPROC = -4

Public Const MOD_CTRL = &H2 'This example uses CTRL
Public Const MOD_SHFT = &H4
Public Const MOD_ALT = &H1
Public Const MOD_EXT = &H8
Public Const MOD_SHIFTCONTROL = &H6
Public Const MOD_ALTSHIFT = &H5
Public Const MOD_CONTROLALT = &H3
Public Const MOD_CONTROLALTSHIFT = &H7

Public glWinRet As Long

Public Const GAME_MINIMIZED = 16
Public Const GAME_UNMINIMIZED = 0

'-----------------------------------

' Function : CallbackMsgs
' This functions is used as a parameter in the
' API SetWindowLong(), by AddresOf operator, so as to
' Subclass the form to get the Windows Callback msgs...
Public Function CallbackMsgs(ByVal wHwnd As Long, ByVal wMsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
    
    'If Message was a hotkey event
    If wMsg = WM_HOTKEY Then
        Call DoFunctions(wp_id)
        CallbackMsgs = 1
        Exit Function
    End If
    
    'If message was from another app aimed at the minmizer
    If wMsg = WM_MINCALLBACK Then
        Select Case wp_id
            Case ID_TOGGLE
                GetCurrentGame
                'If we aren't focused on the game, I guess we're minimized :P
                If GetForegroundWindow() <> gamecon Then
                    MaximizeGame
                Else
                    MinimizeGame (False)
                End If
                CallbackMsgs = 1
                Exit Function
            Case ID_MINIMIZE
                MinimizeGame (False)
                CallbackMsgs = 1
                Exit Function
            Case ID_MAXIMIZE
                MaximizeGame
                CallbackMsgs = 1
                Exit Function
            Case ID_SAVEGAMERES
                SaveGameRes
                CallbackMsgs = 1
                Exit Function
            Case ID_GOTOSCREENRES
                Call ChangeToDesktopResolution(False, True)
                CallbackMsgs = 1
                Exit Function
            Case ID_GOTOGAMERES
                ChangeToGameResolution
                CallbackMsgs = 1
                Exit Function
        End Select
    End If
    
    CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wMsg, wp_id, lp_id)
End Function

' Sub : DoFunction
' Activated by the Function "CallbackMsgs()" whenever
' a hotkey is pressed.
Public Sub DoFunctions(ByVal vKeyID As Byte)

    ' Important Notes :
    ' Do not include any msgboxes or Modal forms in
    ' this procedure, else if you include then by
    ' pressing the Hotkey twice/thrice the application
    ' will be terminated abnormally.
    '
    ' But if it is a requirement for you to include the
    ' Modal forms or msgbox in this procedure then put
    ' the RegisterHotKey() API before hiding the Form
    ' and put the UnRegisterHotKey() API before Showing
    ' the form.

    DoEvents
    ' When the Hotkey is pressed once
    ' check if the Dofunctions() has completed
    ' before the CallbackMsgs().
    ' This check is not required if the form is
    ' minimized in the SysTray ...

    'This is it... This is the code that is activated when the person presses the buttons
    'in order to hide the game.
    If vKeyID = 0 Then
        If OnlyUsingOneHotKey = False Then
            MinimizeGame (False) 'call the minimize game function
        Else
            GetCurrentGame
            'If we aren't focused on the game, I guess we're minimized :P
            If GetForegroundWindow() <> gamecon Then
                MaximizeGame
            Else
                MinimizeGame (False)
            End If
        End If
    ElseIf vKeyID = 1 Then
        MaximizeGame
    End If
End Sub

'www.freevbcode.com - Very good site ^_^
