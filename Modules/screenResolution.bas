Attribute VB_Name = "screenRes"
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
' The following code lists the functions required
' by the code needed to modify the screen resolution.
'
'==================================================

'--------------------------------------------------

'Public ScreenGameWidth, ScreenGameHeight As Long
'Public ScreenMainWidth, ScreenMainHeight As Long

Private Const EWX_REBOOT = 2
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const CDS_FULLSCREEN = &H0
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1

Private Const ENUM_REGISTRY_SETTINGS = -2
Private Const ENUM_CURRENT_SETTINGS = -1


Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private typDM_Screen As typDevMODE
Private typDM_Game   As typDevMODE

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'************************************
'ChangeDisplayResolution
'
'This was the function originally used
'to power the DSM before v1.50.
'I just found it worked by actuaolly re-
'writing the registry, hence DESTROYING
'people's previous settings and mauling
'their desktop images.

'I since built my own functions based off
'of how the Q3 engine changes the res, so it
'be a LOT nicer to use. :)
'************************************
'Public Function ChangeDisplayResolution(NewWidth As Long, NewHeight As Long) As Boolean

'Usage:  ChangeDisplayResolution 800, 600

'Returns: True if succesful, false otherwise

'Comments:  Problems have been reported using this code for
'resolutions higher than 1024 X 768.  We recommend not using this
'snippet to go above this limit.
'Note from TiM: I've tested above that res and things seemed okay.  Don't concern yourself
'too much.


'Dim typDM As typDevMODE
'Dim lRet As Long
'Dim iResp  As Integer

'typDM = pointer to info about current
'display settings

'lRet = EnumDisplaySettings(0, 0, typDM)
'If lRet = 0 Then Exit Function

' Set the new resolution.
'With typDM
'    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
'    .dmPelsWidth = NewWidth
'    .dmPelsHeight = NewHeight
'End With

'Do the update -- Pass update structure to
'ChangeDisplaySettings API function
'lRet = ChangeDisplaySettings(typDM, CDS_UPDATEREGISTRY)
'Select Case lRet
'Case DISP_CHANGE_RESTART

'   iResp = MsgBox _
'  ("You must restart your computer to apply these changes." & _
'        vbCrLf & vbCrLf & "Restart now?", _
'        vbYesNo + vbInformation, "Screen Resolution Changed")
'    If iResp = vbYes Then
'        ChangeDisplayResolution = True
'        Reboot
'    End If

'Case DISP_CHANGE_SUCCESSFUL
'    ChangeDisplayResolution = True
'Case Else
'    ChangeDisplayResolution = False
'End Select

'End Function

'Private Sub Reboot()
'    Dim lRet As Long
'    lRet = ExitWindowsEx(EWX_REBOOT, 0)
'End Sub

Public Sub SaveGameRes()
    Call EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDM_Game)
End Sub

Public Sub ChangeToDesktopResolution(ByVal LostFocus As Boolean, ByVal CalledFromSubClass As Boolean)
    Dim lRet As Long

    GetCurrentGame

    'get resolution settings stored in registry
    Call EnumDisplaySettings(0, ENUM_REGISTRY_SETTINGS, typDM_Screen)

    'If game is currently the foreground (which I'll assume means we're
    'actively using it
    'OR this was called by the onFocus check... in which case, we're already
    'not focusing on the game window >o_O<
    If LostFocus Or (GetForegroundWindow() = gamecon And CalledFromSubClass = False) Then
            'get current resolution
            Call EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDM_Game)
    End If

    'If we actually managed to get a game resolution to change back to after,
    'and the value between the two don't match up.
    If typDM_Game.dmPelsHeight <> 0 And typDM_Game.dmPelsWidth <> 0 And _
        typDM_Screen.dmPelsWidth <> typDM_Game.dmPelsWidth And _
        typDM_Screen.dmPelsHeight <> typDM_Game.dmPelsHeight Then
        
        typDM_Screen.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        
        'The CDS_FULLSCREEN Var tells the computer this res change
        'is temporary, so it leaves desktop icons alone. :)
        lRet = ChangeDisplaySettings(typDM_Screen, CDS_FULLSCREEN)
    End If
End Sub

Public Sub ChangeToGameResolution()
    Dim lRet As Long

    If typDM_Game.dmPelsHeight <> 0 And typDM_Game.dmPelsWidth <> 0 And _
        typDM_Screen.dmPelsWidth <> typDM_Game.dmPelsWidth And _
        typDM_Screen.dmPelsHeight <> typDM_Game.dmPelsHeight Then

        typDM_Game.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        
        lRet = ChangeDisplaySettings(typDM_Game, CDS_FULLSCREEN)
    End If
End Sub
