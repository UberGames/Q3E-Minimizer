VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Q3E Minimizer Main"
   ClientHeight    =   7575
   ClientLeft      =   795
   ClientTop       =   405
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   5228.403
   ScaleMode       =   0  'User
   ScaleWidth      =   4211.647
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFocusMinimize 
      Caption         =   "Minimize Game if it Loses Focus"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   495
      TabIndex        =   5
      Top             =   6360
      Width           =   2595
   End
   Begin VB.TextBox txtAboutScroll 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2895
      Left            =   525
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2025
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CheckBox chkDynaScreenResMod 
      Caption         =   "Dynamic Screen Resolution Modifier"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   495
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   5490
      Width           =   2895
   End
   Begin VB.CheckBox chkRebootCheck 
      Caption         =   "Open this Program on Windows Boot-up"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   495
      TabIndex        =   4
      Top             =   6060
      Width           =   3180
   End
   Begin VB.CheckBox chkShowWinCheck 
      Caption         =   "Hide this Window on Program Start-up"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   495
      TabIndex        =   3
      Top             =   5775
      Width           =   3000
   End
   Begin VB.TextBox txtAbout 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2895
      Left            =   510
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2010
      Width           =   3495
   End
   Begin VB.ComboBox cmbGamesList 
      Enabled         =   0   'False
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   495
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5070
      Width           =   3525
   End
   Begin VB.CommandButton btnOkay 
      Cancel          =   -1  'True
      Caption         =   "Okay!"
      Default         =   -1  'True
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   6855
      Width           =   3525
   End
End
Attribute VB_Name = "frmMain"
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
' This is the main form that appears upon first
' executing Q3E Minimizer. It contains a brief
' introduction, as well most of the major settings
' for the program.
'==================================================

Option Explicit

Dim rgnBasic As New Region
Dim rgnControl(3) As New Region
Dim CurrentRgn As Long
Dim pic As New StdPicture

Private Sub btnOkay_Click()
    Me.Hide 'Used to be Unload Me but that screwed up the system tray a tad
            'FIXME: Now is global so shoudln't matter and could possibly save memory
            'UNFIXME: Okay... not global enough... whole program exploded O_O
End Sub

Private Sub chkDynaScreenResMod_Click()
    Call WriteSettings("DynaScreen", chkDynaScreenResMod.Value)
End Sub

Private Sub chkFocusMinimize_Click()
    Call WriteSettings("MinimizeOnUnFocus", chkFocusMinimize.Value)
End Sub

Private Sub chkRebootCheck_Click()
    Call WriteSettings("StartOnBoot", chkRebootCheck.Value)
    
    'Write startup data to registry
    If chkRebootCheck.Value = 1 Then
        RunAtStartup App.Title, App.Path & "\" & App.EXEName & ".EXE"
    Else
        RemoveFromStartup App.Title, App.Path & "\" & App.EXEName & ".EXE"
    End If
End Sub

Private Sub cmbGamesList_Click()
    Call WriteSettings("CurrentGame", cmbGamesList.ListIndex)
    DoIconz
End Sub

Private Sub Form_Load()
    Dim i As Integer
    txtAbout.Text = "Q3E Minimizer is a small program that will allow you to minimize nearly any Q3 based game anytime, in-game without having to suffer any pesky lag or loading times." + vbNewLine + "To customize this program, right-click on the Q3 Logo icon in the Windows System tray and choose 'Settings'." + vbNewLine + vbNewLine + "To close this program, right-click on it's icon in the system tray and choose 'Exit'.  (The game will automatically restore if it's open.)" + vbNewLine + vbNewLine + "Happy Fragging!!" + vbNewLine + "-TiM"

    'Thanks to Brian Yule for the brillant code to make windows' edges transparent!
    'Borrowed from the tips site of http://www.vbrad.com!
    ' Load pictures from file
    Set pic = Me.picture
    ' Scan Shape from Green Screen Style Image
    Call rgnBasic.ScanPicture(pic)
    ' Offset the Shape to allow for the form header.
    'Call rgnBasic.OffsetHeader(Me)

    Me.picture = pic ' Set the Form Background
    Call rgnBasic.ApplyRgn(Me.hWnd) ' Set the Form Shape
    CurrentRgn = rgnBasic.hndRegion ' Set the Current Shape
    
    'Populate Listbox from games list
    For i = 0 To (MAX_GAMES - 1)
        cmbGamesList.AddItem games(i).gameFormal, i
    Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Thanks to the MSDN database site for the code that lets you shift
    'a window when clicking on an object.
    Dim lngReturnValue As Long

    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub chkShowWinCheck_Click()
    Call WriteSettings("HideWinOnStart", chkShowWinCheck.Value)
End Sub
