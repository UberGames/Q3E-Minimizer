VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Region"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'==================================================
' Q3E Minimizer 1.51 by UberGames
' Developed by Timothy 'TiM' Oliver
'
' This code was created by Brian Yule in order to
' set up the Region class applied to the background
' images to set up the transparency.
'==================================================

Option Explicit
'Thanks Brian Yule!!
Public hndRegion As Long
Private DIB As New cDIBSection
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub Class_Terminate()
    If hndRegion <> 0 Then Call DeleteObject(hndRegion)
End Sub

Public Sub ApplyRgn(ByVal hWnd As Long)
'
'   When the setWinowRgn function gets a handle to a region it applies and
'   deletes the region.  Therefore you cannot call the function twice with
'   the handle to the same region.  Therefore I make a copy of the region
'   for windows to apply and delete.
'
    Dim hndRegionCopy As Long
    
    hndRegionCopy = CreateRectRgn(0, 0, 0, 0) ' Create a blank region

    Call CombineRgn(hndRegionCopy, hndRegion, hndRegionCopy, RGN_OR) ' Copy the region
    Call SetWindowRgn(hWnd, hndRegionCopy, True)
End Sub

Public Sub ScanPicture(ByVal picture As StdPicture, Optional transColor As Long = vbNull)
    Dim Rgn2 As Long
    
    hndRegion = CreateRectRgn(0, 0, 0, 0)
    
    Dim x As Long, y As Long
    Dim SPos As Long, EPos As Long
    Dim Wid As Long, Hgt As Long
    Dim bDib() As Byte
    Dim tSA As SAFEARRAY2D
      
    'get the picture size of the form
    Call DIB.CreateFromPicture(picture)
    
    Wid = DIB.Width
    Hgt = DIB.Height
    
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.BytesPerScanLine
        .pvData = DIB.DIBSectionBitsPtr
    End With
    
    Call CopyMemory(ByVal VarPtrArray(bDib), VarPtr(tSA), 4)
    
    ' if there is no transColor specified, use the first pixel as the transparent color
    If transColor = vbNull Then transColor = RGB(bDib(0, 0), bDib(1, 0), bDib(2, 0))
    
    For y = 0 To DIB.Height - 1 'line scan
        x = -3
        Do
            Rgn2 = 0
            x = x + 3
            
            While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) = transColor And (x < DIB.Width * 3 - 3)
                x = x + 3 'skip the transparent point
            Wend
            
            SPos = x / 3
            
            While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) <> transColor And (x < DIB.Width * 3 - 3)
                x = x + 3 'skip the nontransparent point
            Wend
            
            EPos = x / 3
            
            'combine the region
            If SPos <= EPos Then
                Rgn2 = CreateRectRgn(SPos, DIB.Height - y, EPos, DIB.Height - 1 - y)
                Call CombineRgn(hndRegion, hndRegion, Rgn2, RGN_OR)
                Call DeleteObject(Rgn2)
            End If
        Loop Until x >= DIB.Width * 3 - 3
    Next y
    
    Call CopyMemory(ByVal VarPtrArray(bDib), 0&, 4)
End Sub

Public Sub OffsetHeader(ByRef tmpForm As Form)
    With tmpForm
        If .BorderStyle <> 0 Then
            Dim xoff As Long, yoff As Long
            
            .ScaleMode = vbPixels
            
            xoff = (.ScaleX(.Width, vbTwips, vbPixels) - .ScaleWidth) / 2
            yoff = .ScaleY(.Height, vbTwips, vbPixels) - .ScaleHeight - xoff
    
            Call OffsetRgn(hndRegion, xoff, yoff)
        End If
    End With
End Sub
