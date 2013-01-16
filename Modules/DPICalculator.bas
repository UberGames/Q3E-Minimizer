Attribute VB_Name = "DPICalculator"
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
' The following code is used to determine the user's
' current screen DPI settings (Which can be custom
' controlled in the Windows Control Panel), so that
' in the event the characters are extrmemely large,
' the program can adapt the windows and font size
' so the program GUI isn't distorted.
'==================================================

'--------------------------------------------------

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Function GetDPIX() As Long
    Dim hDC As Long

    hDC = GetDC(0)
    GetDPIX = GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
End Function

Public Function GetDPIY() As Long
    Dim hDC As Long

    hDC = GetDC(0)
    GetDPIY = GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC
End Function

Public Function EstimateFontSize(ByVal sngPointSize As Single) As Single
    Const baseDPI As Long = 96

    EstimateFontSize = sngPointSize * baseDPI / GetDPIY
End Function
