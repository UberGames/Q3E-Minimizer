VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Q3E Minimizer Settings"
   ClientHeight    =   2130
   ClientLeft      =   6195
   ClientTop       =   6405
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSettings.frx":0000
   ScaleHeight     =   2130
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInstructionsScroll 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   1080
      Left            =   525
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   1080
      Left            =   525
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton btnOkay 
      Caption         =   "Okay!"
      Default         =   -1  'True
      Height          =   240
      Left            =   4230
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox txtRestoreKeyz 
      Height          =   300
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtHideKeyz 
      Height          =   300
      Left            =   4965
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3120
   End
   Begin VB.Label lblRstore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Restore:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4245
      TabIndex        =   6
      Top             =   1005
      Width           =   840
   End
   Begin VB.Label lblHide 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Hide:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4245
      TabIndex        =   5
      Top             =   630
      Width           =   615
   End
End
Attribute VB_Name = "frmSettings"
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
' This is the settings form that allows you to set
' new hotkeys for the program to use.
'==================================================

Option Explicit

Dim rgnBasic As New Region
Dim CurrentRgn As Long
Dim pic As New StdPicture

'Used so the user can't get away with just
'using the control/alt/shift buttons alone,
'coz that wouldn't be good
Dim KeysAreValid As Boolean

Public Function GetKeySequence(ByVal KeyCode As Integer, ByVal Shift As Integer) As String
    Dim strTextCombination As String
    
    If (KeyCode = 0) Then
        GetKeySequence = "None"
        Exit Function
    End If
    
    KeysAreValid = False

    'Thanks to the fine help on experts-exchange.com for this little snippet of code. :)
    'I can't remember who it was exactly who created this.  So if that person is you,
    'please tell me and I'll update credit accrdingly. :)

    If (Shift And vbCtrlMask) Then strTextCombination = strTextCombination & "+Control"
    If (Shift And vbAltMask) Then strTextCombination = strTextCombination & "+Alt"
    If (Shift And vbShiftMask) Then strTextCombination = strTextCombination & "+Shift"

    Select Case KeyCode
        Case vbKey0 To vbKey9
            strTextCombination = strTextCombination & "+" & Chr(KeyCode)
            KeysAreValid = True
        Case vbKeyA To vbKeyZ
            strTextCombination = strTextCombination & "+" & Chr(KeyCode)
            KeysAreValid = True
    End Select
    
    'The user pressed something that wasn't up there O_o
    'Must be one of the other, special keys
    If KeysAreValid = False Then
        Dim specialChar As String
        specialChar = ASCIIChars(KeyCode)
        
        'If we got a result lol
        If Len(specialChar) Then
            strTextCombination = strTextCombination & "+" & specialChar
            KeysAreValid = True
        End If
    End If

    GetKeySequence = Mid(strTextCombination, 2)

End Function

Private Sub btnOkay_Click()

    'If user tried to get away with not setting a hot key.
    If txtHideKeyz.Text = "None" Then
        hotKeysData(HK_MINIMIZE).KeyASCII = 0
        hotKeysData(HK_MINIMIZE).ShiftFlags = 0
        'MsgBox ("Error: Please enter something into the Hot Key field to Hide the game!"), vbCritical
        'Exit Sub
    End If
    
    If txtRestoreKeyz.Text = "None" Then
        hotKeysData(HK_RESTORE).KeyASCII = 0
        hotKeysData(HK_RESTORE).ShiftFlags = 0
        'MsgBox ("Error: Please enter something into the Hot Key field to Restore the game!"), vbCritical
        'Exit Sub
    End If

    'If user tried to assign the same keys to both
    'If txtRestoreKeyz.Text = txtHideKeyz.Text Then
    '    MsgBox ("Error: You can't assign the same key commands to both show and hide the game! Bad! Smite! Evil!"), vbCritical
    '    Exit Sub
    'End If

    'Write settings to registry
    Call WriteSettings("HideQ3Keys", hotKeysData(HK_MINIMIZE).ShiftFlags & " " & hotKeysData(HK_MINIMIZE).KeyASCII)
    Call WriteSettings("ShowQ3Keys", hotKeysData(HK_RESTORE).ShiftFlags & " " & hotKeysData(HK_RESTORE).KeyASCII)

    'Unregister current hotkeys
    UnregisterHotKey frmSubliminal.hWnd, HK_MINIMIZE
    UnregisterHotKey frmSubliminal.hWnd, HK_RESTORE

    'Just in case some weird users want to use the same key to toggle the screen
    If hotKeysData(HK_MINIMIZE).KeyASCII = hotKeysData(HK_RESTORE).KeyASCII And _
        hotKeysData(HK_MINIMIZE).ShiftFlags = hotKeysData(HK_RESTORE).ShiftFlags Then
        OnlyUsingOneHotKey = True
    Else
        OnlyUsingOneHotKey = False
    End If

    'Make this false, so the timer module will look at this
    'and then re-register the keys if need be
    'If TriedToRegister Then
    '    TriedToRegister = False
    'End If

    'Go ahead and register the hotkeys.  Before, the timer
    'would delay the hotkey registration.  Which got annoying in some
    'cases.
    GetCurrentGame
    If gamecon <> 0 Then
        retVal0 = RegisterHotKeys(HK_MINIMIZE)
        If OnlyUsingOneHotKey = False Then
            retVal1 = RegisterHotKeys(HK_RESTORE)
        End If
        TriedToRegister = True
    Else
        TriedToRegister = False
    End If
    

    frmSettings.Hide
End Sub

Private Sub Form_Load()
    ' Load pictures from file
    Set pic = Me.picture
    ' Scan Shape from Green Screen Style Image
    Call rgnBasic.ScanPicture(pic)
    ' Offset the Shape to allow for the form header.
    'Call rgnBasic.OffsetHeader(Me)
    
    Me.picture = pic ' Set the Form Background
    Call rgnBasic.ApplyRgn(Me.hWnd) ' Set the Form Shape
    CurrentRgn = rgnBasic.hndRegion ' Set the Current Shape
    
    txtHideKeyz.Text = GetKeySequence(hotKeysData(HK_MINIMIZE).KeyASCII, hotKeysData(HK_MINIMIZE).ShiftFlags)
    txtRestoreKeyz.Text = GetKeySequence(hotKeysData(HK_RESTORE).KeyASCII, hotKeysData(HK_RESTORE).ShiftFlags)
    
    txtInstructions.Text = "Click on a text field and then press the key combination you want.  The key sequence can be any of the Shift,Alt, and/or Ctrl keys plus a normal key. You are able to use the same key sequence for both fields."
    txtInstructionsScroll.Text = txtInstructions.Text

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngReturnValue As Long

    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub txtHideKeyz_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeysAreValid = False Then
        txtHideKeyz.Text = GetKeySequence(KeyCode, Shift)
    End If
    
    If txtHideKeyz.Text = "" Then
        txtHideKeyz.Text = "None"
    End If
End Sub

Private Sub txtRestoreKeyz_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeysAreValid = False Then
        txtRestoreKeyz.Text = GetKeySequence(KeyCode, Shift)
    End If
    
    If txtRestoreKeyz.Text = "" Then
        txtRestoreKeyz.Text = "None"
    End If
End Sub

Private Sub txtHideKeyz_KeyDown(KeyCode As Integer, Shift As Integer)
    'Unregister the key, or weird stuff may happen
    GetCurrentGame
    
    'Bug fix: Unregister both. In case the user tries to swap keys etc :P
    If gamecon <> 0 Then
        UnregisterHotKey frmSubliminal.hWnd, HK_MINIMIZE
        UnregisterHotKey frmSubliminal.hWnd, HK_RESTORE
    End If
    
    txtHideKeyz.Text = GetKeySequence(KeyCode, Shift)

    If KeysAreValid Then
        hotKeysData(HK_MINIMIZE).ShiftFlags = Shift
        hotKeysData(HK_MINIMIZE).KeyASCII = KeyCode
    End If
    
    If txtHideKeyz.Text = "" Then
        txtHideKeyz.Text = "None"
    End If
End Sub

Private Sub txtRestoreKeyz_KeyDown(KeyCode As Integer, Shift As Integer)
     'Unregister the key, or weird stuff may happen
    GetCurrentGame
    
    If gamecon <> 0 Then
        UnregisterHotKey frmSubliminal.hWnd, HK_MINIMIZE
        UnregisterHotKey frmSubliminal.hWnd, HK_RESTORE
    End If
    
    txtRestoreKeyz.Text = GetKeySequence(KeyCode, Shift)
         
    If KeysAreValid Then
        hotKeysData(HK_RESTORE).ShiftFlags = Shift
        hotKeysData(HK_RESTORE).KeyASCII = KeyCode
    End If
    
    If txtRestoreKeyz.Text = "" Then
        txtRestoreKeyz.Text = "None"
    End If
End Sub
