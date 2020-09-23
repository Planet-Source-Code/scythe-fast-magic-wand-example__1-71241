Attribute VB_Name = "modRegionShape2"
Option Explicit
' BRIEF HISTORY
' \/\/\/\/\/\/\
' *** VERSION 1 - used GetPixel
' This approach should be somewhat faster than typical approaches that
' create a region by using CreateRectRgn, CombineRgn & DeleteObject.
' That is 'cause this approach does not use any of those functions to
' create regions. Like typical approaches a rectangle format is used to
' "add to a region", however, this approach directly creates the region header
' & region structure and passes that to a single API to finish the job of
' creating the end result. For those that play around with such things,
' I think you will recognize the difference.

' *** VERSION 2 - use of DIBs vs GetPixel
' UPDATED to include using DIB sections. The increase in speed is
' truly amazing over my already faster earlier version.....
' Over a 300% increase in speed noted on smallish bitmaps (96x96) &
' over a 425% increase noted on mid-sized bitmaps (265x265)

' EDITED: The ExtCreateRegion seems to have an undocumented restriction: it
' won't create regions comprising of more than 4K rectangles (win98). So to
' get around this for very complex bitmaps, I rewrote the function to create
' regions of 2K rectangles at a time if needed. This is still extremely
' fast. I compared the window shaping code from vbAccelerator with the
' SandStone.bmp (15,000 rects) in Windows directory. vbAccelerator's code
' averaged 4,900 ms. My routines averaged 77 ms & that's not a typo!
' EDITED: Using CopyImage on a form's picture property screwed up NT4.
'  Why? Have no idea. But replaced that line of code.

' EDITED: I was allowing default error trapping on UBound() to resize the
' rectangle array: when trying to update array element beyond UBound, error
' would occur & be redirected to resize the array. However, thanx to
' Robert Rayment, if the UBound checks are disabled in compile optimizations,
' then we get a crash. Therefore, checks made appropriately & a small loss of
' speed is the trade-off for safety.

' *** VERSION 2.1 - Use of GetDIBbits vs SafeArray pointers
' The function CreateShapedRegion2 does not use all that fancy SafeArray stuff
' It was tweaked to use the bytes returned by GetDIBbits() API
' Otherwise it is basically the same as Version 2
' ***********************************************************************

' GDI32 APIs
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal HDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal HDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
' Kernel32 APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
' User32 APIs
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' TYPEs used
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
' Constants used
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const RGN_OR As Long = 2

' for testing only
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function CreateShapedRegion2(ByVal hBitmap As Long, ByVal hwnd As Long, Optional ByVal transColor As Long = -1) As Long
' hBitmap = handle to a bitmap to be used to create the region
' hWnd = handle to the window that will be shaped

' test for required variables first
If hBitmap = 0 Or hwnd = 0 Then Exit Function
' now ensure hBitmap handle passed is a usable bitmap
Dim bmpInfo As BITMAPINFO
If GetGDIObject(hBitmap, Len(bmpInfo), bmpInfo) = 0 Then Exit Function

' declare bunch of variables...
Dim srcDC As Long, tDC As Long, dDC As Long
Dim oldBmp As Long, oldDIB As Long

Dim rgnRects() As RECT, rectCount As Long
Dim rStart As Long, bReset As Boolean, lRgnCount As Long

Dim X As Long, Y As Long
Dim lWidth As Long, lHeight As Long, lScanLines As Long

Dim bDib() As Byte, hDib As Long
Dim bBGR(0 To 3) As Byte, tgtColor As Long
Dim tRgn As Long, xRgn As Long

Dim aStart As Long, aStop As Long ' testing purposes
aStart = GetTickCount()            ' testing purposes

On Error GoTo CleanUp
' create a temporary DC to hold the bitmap copy & another to hold new DIB
srcDC = GetDC(hwnd)
tDC = CreateCompatibleDC(srcDC)     ' for source bitmap
ReleaseDC hwnd, srcDC
' select bitmap copy into the DC & get default transparent color if needed
oldBmp = SelectObject(tDC, hBitmap)
lWidth = bmpInfo.bmiHeader.biWidth
lHeight = bmpInfo.bmiHeader.biHeight
   
' Scans must align on dword boundaries:
lScanLines = (lWidth * 3 + 3) And &HFFFFFFFC
ReDim bDib(0 To lScanLines - 1, 0 To lHeight - 1)
With bmpInfo.bmiHeader
   .biSize = Len(bmpInfo.bmiHeader) ' Set up for image
   .biBitCount = 24
   .biHeight = -.biHeight
   .biPlanes = 1
   .biCompression = BI_RGB
   .biSizeImage = lScanLines * .biHeight
End With
' create DIB to prevent any bitCount mismatches: forces to 24bpp
hDib = CreateDIBSection(tDC, bmpInfo, DIB_RGB_COLORS, ByVal 0&, 0, 0)
If hDib Then
    ' select DIB into DC & then blt source image on it
    dDC = CreateCompatibleDC(tDC)     ' for dib
    oldDIB = SelectObject(dDC, hDib)
    BitBlt dDC, 0, 0, lWidth, lHeight, tDC, 0, 0, vbSrcCopy
    ' get the image out of DC; not needed inside anymore
    SelectObject tDC, oldBmp
    DeleteDC tDC
    tDC = 0

    ' now get the transparent color if needed
    If transColor < 0 Then transColor = GetPixel(dDC, 0, 0)
    If transColor > -1 Then
        ' my tweak: DIB colors are BGR vs RGB, so we will get the
        ' RGB values and save to a long value as BGR for faster comparisons later
        bBGR(2) = (transColor And &HFF&)
        bBGR(1) = (transColor And &HFF00&) \ &H100&
        bBGR(0) = (transColor And &HFF0000) \ &H10000
        bBGR(3) = 0: transColor = 0
        CopyMemory transColor, bBGR(0), 4&
        
        ' get the DIB out of DC; not needed inside anymore & return byte array
        SelectObject dDC, oldDIB
        Call GetDIBits(dDC, hDib, 0, lHeight, bDib(0, 0), bmpInfo, 0)
        DeleteDC dDC
        dDC = 0
            
         ' start with an arbritray number of rectangles
         ReDim rgnRects(0 To lWidth * 3)
         ' DIB images are upside down; need to account for that when building
         ' regional rectangles; which must be left to right, top to bottom
         tgtColor = 0

         For Y = 0 To lHeight - 1
             For X = 0 To lWidth - 1
                 ' my hack: we already saved a long as BGR, now
                 ' get the current DIB pixel into a long (BGR also) & compare
                 CopyMemory tgtColor, bDib(X * 3, Y), 3&
                 If transColor <> tgtColor Then
                     If bReset Then
                         ' we're currently tracking a rectangle, so let's close it
                        If rectCount + 1 = UBound(rgnRects) Then _
                            ReDim Preserve rgnRects(0 To UBound(rgnRects) + lWidth * 3)
                         SetRect rgnRects(rectCount + 2), rStart, Y, X, Y + 1
                         bReset = False          ' reset flag
                         rectCount = rectCount + 1     ' keep track of nr in use
                     End If
                 Else
                     ' not a transparent color
                     If bReset = False Then
                         ' set flag to indicate tracking non-transparent pixels
                         bReset = True
                         rStart = X      ' set start point
                     End If
                 End If
             Next X
             If bReset Then
                 ' got to end of bitmap without hitting another transparent pixel
                 ' but we're tracking so we'll close rectangle now
                If rectCount + 1 = UBound(rgnRects) Then _
                    ReDim Preserve rgnRects(0 To UBound(rgnRects) + lWidth * 3)
                 SetRect rgnRects(rectCount + 2), rStart, Y, X, Y + 1
                 bReset = False          ' reset flag
                 rectCount = rectCount + 1     ' keep track of nr in use
             End If
        Next Y
        Erase bDib
        On Error Resume Next
        If rectCount Then
            tRgn = CreatePartialRegion(rgnRects(), 2, rectCount + 1, bmpInfo.bmiHeader.biWidth)
            ' ok, now to test whether or not we are good to go...
            If tRgn = 0 And rectCount > 2000 Then
                ' Win98 has problems with regional rectangles over 4000
                ' So, we'll try again in case this is the prob with other
                ' systems too. This time we'll step it at 2000 at a time
                For X = 2 To rectCount + 1 Step 2000
                    If X + 2000 > rectCount + 1 Then
                        Y = rectCount + 1
                    Else
                        Y = X + 2000
                    End If
                    xRgn = CreatePartialRegion(rgnRects(), X, Y, bmpInfo.bmiHeader.biWidth)
                    If xRgn = 0 Then
                        If tRgn Then DeleteObject tRgn
                        tRgn = 0
                        Exit For
                    End If
                    If tRgn Then
                        ' use combineRgn, but only every 2000th time
                        CombineRgn tRgn, tRgn, xRgn, RGN_OR
                        DeleteObject xRgn
                    Else
                        tRgn = xRgn
                    End If
                Next
            End If
        End If
        'SetWindowRgn hwnd, tRgn, True
        If tRgn Then
            aStop = GetTickCount()          ' testing purposes
            'CreateShapedRegion2 = True
            CreateShapedRegion2 = tRgn
        Else
            MsgBox "Shaped Region failed. Windows could not create the region."
        End If
    Else
        'SetWindowRgn hwnd, ByVal 0&, True
        MsgBox "Invalid transparent color value at location 0,0 in source image"
    End If
Else
    'SetWindowRgn hwnd, ByVal 0&, True
    MsgBox "Shaped Region failed. Failed to create DIB section."
End If


CleanUp:
Erase rgnRects()
If tDC Then
    SelectObject tDC, oldBmp
    DeleteDC tDC
End If
If dDC Then
    SelectObject dDC, oldDIB
    DeleteDC dDC
End If

If hDib Then DeleteObject hDib
If Err Then Err.Clear
' testing purposes
'If aStop Then MsgBox aStop - aStart & " ms"

End Function

Private Function CreatePartialRegion(rgnRects() As RECT, lIndex As Long, uIndex As Long, cX As Long) As Long
On Error Resume Next
' Note: Ideally contiguous rectangles should be combined into one larger
' rectangle. However, thru trial & error I found that Windows does this
' for us and taking the extra time to do it ourselves slows down the results.

' the first 32 bytes of a region describe the region.
' Well 32 bytes equates to 2 rectangles (16bytes each), so
' I'll cheat & use rectangles to store the data
With rgnRects(lIndex - 2) ' bytes 0-15
    .Left = 32                      ' length of region header in bytes
    .Top = 1                        ' required cannot be anything else
    .Right = uIndex - lIndex + 1    ' number of rectangles for the region
    .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
End With
With rgnRects(lIndex - 1) ' bytes 16-31 bounding rectangle identification
    .Left = 0                           ' left
    .Top = rgnRects(lIndex).Top         ' top
    .Right = cX                         ' right
    .Bottom = rgnRects(uIndex).Bottom   ' bottom
End With
' call function to create region from our byte (RECT) array
CreatePartialRegion = ExtCreateRegion(ByVal 0&, (rgnRects(lIndex - 2).Right + 2) * 16, rgnRects(lIndex - 2))
If Err Then Err.Clear
End Function

