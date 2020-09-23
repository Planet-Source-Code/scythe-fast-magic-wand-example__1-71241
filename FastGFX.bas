Attribute VB_Name = "FastGFX"
Option Explicit

Public Declare Function FloodFill Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal HDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type


Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

Private Const DIB_RGB_COLORS As Long = 0

Public Buf1() As RGBQUAD
Public Buf2() As RGBQUAD


'Convert Picture to Array
Public Sub Pic2Array(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)

 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
 ReDim PicArray(0 To PicBox.ScaleWidth - 1, 0 To PicBox.ScaleHeight - 1)

With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
End With
'Now get the Picture
GetDIBits PicBox.HDC, PicBox.Image.handle, 0, Binfo.bmiHeader.biHeight, PicArray(0, 0), Binfo, DIB_RGB_COLORS

End Sub
'Convert Array to Picture
Public Sub Array2Pic(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)

 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos

With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
End With
SetDIBits PicBox.HDC, PicBox.Image.handle, 0, Binfo.bmiHeader.biHeight, PicArray(0, 0), Binfo, DIB_RGB_COLORS

End Sub
'Monochrome a Picture
Public Sub PicMonochrome(PicAr() As RGBQUAD)
 Dim X As Long
 Dim Y As Long
 Dim col As Long

 For X = 0 To UBound(PicAr, 1)
  For Y = 0 To UBound(PicAr, 2)
   'calculate the Colors
   'Red * 0,3 + Green * 0,59 + Blue * 0,11 gives us the Graycolor
   'The Maximum result is 255
   col = 0.3 * CLng(PicAr(X, Y).rgbRed) + 0.59 * CLng(PicAr(X, Y).rgbGreen) + 0.11 * CLng(PicAr(X, Y).rgbBlue)
   'For this we only need the red channel
   PicAr(X, Y).rgbRed = col
   'PicAr(x, y).rgbGreen = col
   'PicAr(x, y).rgbBlue = col
  Next Y
 Next X
End Sub
'Hue a Picture
'Not my code
'found it somewhere in the www
'needs to be optimized
Public Sub PicHue(PicAr() As RGBQUAD)
 Dim X As Long
 Dim Y As Long
 Dim R As Integer, G As Integer, B As Integer
 Dim cMax As Integer, cMin As Integer
 Dim RDelta As Double, GDelta As Double, BDelta As Double
 Dim H As Single
 Dim s As Single
 Dim l As Single
 Dim cMinus As Long, cPlus As Long
 Dim notthere As Boolean

 For X = 0 To UBound(PicAr, 1)
  For Y = 0 To UBound(PicAr, 2)
   R = PicAr(X, Y).rgbRed
   G = PicAr(X, Y).rgbGreen
   B = PicAr(X, Y).rgbBlue
   
   'Calculate the hue
   cMax = Maximum(R, G, B) 'iMax(iMax(R, G), B) 'Highest and lowest
   cMin = Minimum(R, G, B) 'iMin(iMin(R, G), B) 'color values

   cMinus = cMax - cMin 'Used to simplify the
   cPlus = cMax + cMin  'calculations somewhat.
   If cMax = cMin Then 'achromatic (r=g=b, greyscale)
    H = 160
   Else
    RDelta = ((cMax - R) * 40 + 0.5) / cMinus
    GDelta = ((cMax - G) * 40 + 0.5) / cMinus
    BDelta = ((cMax - B) * 40 + 0.5) / cMinus

 If cMax = CLng(R) Then
     H = BDelta - GDelta
 ElseIf cMax = CLng(G) Then
     H = 80 + RDelta - BDelta
 Else
     H = 160 + GDelta - RDelta
 End If


    If H < 0 Then H = H + 240
   End If

   'For this we only need the red channel
   PicAr(X, Y).rgbRed = H
   'PicAr(x, y).rgbGreen = H
   'PicAr(x, y).rgbBlue = H
  Next Y
 Next X

End Sub
Public Function Maximum(rR As Integer, rG As Integer, rB As Integer) As Integer
    If (rR > rG) Then
        If (rR > rB) Then Maximum = rR Else Maximum = rB
      Else
        If (rB > rG) Then Maximum = rB Else Maximum = rG
    End If
End Function

Public Function Minimum(rR As Integer, rG As Integer, rB As Integer) As Integer
    If (rR < rG) Then
        If (rR < rB) Then Minimum = rR Else Minimum = rB
      Else
        If (rB < rG) Then Minimum = rB Else Minimum = rG
    End If
End Function

'Check if a Color is ind a range X% from the actual point
Public Function SimilarColor(ByVal Red1 As Long, ByVal Green1 As Long, ByVal Blue1 As Long, ByVal Red2 As Long, ByVal Green2 As Long, ByVal Blue2 As Long, ByVal Percent As Long) As Boolean
 'We have 255 Colors so wen need 100*2.55 to get all
 Percent = Percent * 2.55
 'Check if the color is in our range
 If Abs(Red1 - Red2) <= Percent And Abs(Green1 - Green2) <= Percent And Abs(Blue1 - Blue2) <= Percent Then SimilarColor = True
End Function

Public Function SameColor(Red As Byte, Blue As Byte, Green As Byte) As Byte
Dim Tmp As Byte
If Red > Green Then Tmp = 1
If Green > Blue Then Tmp = Tmp + 10
If Red > Blue Then Tmp = Tmp + 100
SameColor = Tmp
End Function
