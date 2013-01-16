Attribute VB_Name = "controlRegion"
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
'Thanks to Garrett Sever (aka "The Hand") for this
'section of code.
'This section of code declares the functions needed
'to make the checkboxes in frmMain transparent around
'the edges, making it look much more pretty. ;)
'==================================================

'Private Const RGN_AND = 1   'Creates the intersection of the two combined regions.
'Private Const RGN_COPY = 5  'Creates a copy of the region identified by hrgnSrc1.
'Private Const RGN_OR = 2    'Creates the union of two combined regions.
Private Const RGN_XOR = 3   'Creates the union of two combined regions except for any overlapping areas.
'Private Const RGN_DIFF = 4  'Combines the parts of hrgnSrc1 that are not part of hrgnSrc2.

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Declares for color retrieval
'*****************************************
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Const COLOR_BTNFACE = 15 'Button
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub TransChkOpt(aCtl As Control, Form As Form, trans As Boolean)

    Dim x               As Long     'Used to loop thru the pixels
    Dim y               As Long     'Used to loop thru the pixels
    Dim Wid             As Long     'Width of optButton - Used to loop thru the pixels
    Dim Hgt             As Long     'Height of optButton - Used to loop thru the pixels

    Dim rgnCtl          As Long     'Region of the Button
    Dim rgnPixel        As Long     'Region of a pixel - used to subtract out tiny areas
    Dim colPixel        As Long     'Color of a pixel in the control's DC
    Dim ctlDC           As Long     'Temporary device context used to get pixel color info
    Dim backColor       As Long     'Background color of the option button
    
    'Trans color for the button
    backColor = GetSysColor(COLOR_BTNFACE)
    
    
    'Calculate the size of the picture which we will fit the form to
    Wid = Form.ScaleX(aCtl.Width, Form.ScaleMode, vbPixels)
    Hgt = Form.ScaleX(aCtl.Width, Form.ScaleMode, vbPixels)

    'Create a region the same size as our picture dimensions
    rgnCtl = CreateRectRgn(0, 0, Wid, Hgt)
    If Not trans Then GoTo TransChkOpt_SetRgn
    
    'Capture the option button's DC so we can read the color information.
    ctlDC = GetDC(aCtl.hWnd)
    
    'Loop thru all pixels in the option button
    For y = 0 To Hgt
        For x = 0 To Wid
            ' check the color of each pixel
            colPixel = GetPixel(ctlDC, x, y)
            If colPixel = backColor Then
                'If the color is our mask color (button face in this case) then
                ' create a tiny region for it and remove it from the picture
                rgnPixel = CreateRectRgn(x, y, x + 1, y + 1)
                CombineRgn rgnCtl, rgnCtl, rgnPixel, RGN_XOR
                'Clean up our graphics resource

                DeleteObject rgnPixel
            End If
        Next x
    Next y
            
    Call SetWindowRgn(aCtl.hWnd, rgnCtl, True)
    
    'Release our control's DC
    ReleaseDC aCtl.hWnd, ctlDC
    Exit Sub
    
TransChkOpt_SetRgn:
    'Fit the button to the new region.
    DeleteObject SetWindowRgn(aCtl.hWnd, rgnCtl, True)
'Thanks for the cool code and for making it available online Garrett!
End Sub
