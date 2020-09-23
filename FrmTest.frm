VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "Testform Fast Magic Wand Selection"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CboStyle 
      Height          =   315
      ItemData        =   "FrmTest.frx":0000
      Left            =   1920
      List            =   "FrmTest.frx":0010
      Style           =   2  'Dropdown-Liste
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton OptPic 
      Caption         =   "Border"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton OptPic 
      Caption         =   "Magic"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton OptPic 
      Caption         =   "Picture"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Timer TmrShowBorder 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7560
      Top             =   240
   End
   Begin VB.HScrollBar HscrPerc 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   480
      Value           =   15
      Width           =   3375
   End
   Begin VB.PictureBox PicOrg 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   6750
      Left            =   120
      OLEDropMode     =   1  'Manuell
      Picture         =   "FrmTest.frx":0031
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   9000
   End
   Begin VB.PictureBox PicBorder 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   6750
      Left            =   120
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   9000
   End
   Begin VB.PictureBox PicDest 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   6750
      Left            =   120
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   9000
   End
   Begin VB.Label Label2 
      Caption         =   "Selected Color"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label LblTime 
      Caption         =   "  "
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Shape ShpColor 
      FillStyle       =   0  'Ausgefüllt
      Height          =   255
      Left            =   5400
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Difference 15%"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Real Fast Magic Wand
'Scythe 2008

'How does ist work
'I take a color and clear anything thats different

'After this i use FloodFill to make the everything
'white thats touching our original point

'Now i create a line on every white pixel with a non white near it

'Thats it
'Very simple but extreme fast :-)

'Pointer to our Region
Dim MasterRgn As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

'To show the border
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const PATINVERT = &H5A0049        ' dest = pattern XOR dest
Dim OrOrNot As Boolean



Private Sub Form_Load()
 CboStyle.ListIndex = 0
End Sub

'Show the allowed difference in %
Private Sub HscrPerc_Change()
 Label1 = "Difference  " & HscrPerc.Value & "%"
End Sub

'Show the different Pictures
'You can make this with virtual Pic´s
'but i wantet to see and show it
Private Sub OptPic_Click(Index As Integer)
 PicOrg.Visible = OptPic(0).Value
 PicDest.Visible = OptPic(1).Value
 PicBorder.Visible = OptPic(2).Value
End Sub


'The Mainroutine
Private Sub PicOrg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Xl As Long          'X position
 Dim Yl As Long          'Y position
 Dim Ro As Byte          'Red Origin
 Dim Go As Byte          'Green Origin
 Dim Bo As Byte          'Blue Origin
 Dim Percent As Long     'Allowed RGB Difference in %
 Dim Tmr As Long         'To test how long it takes
 Dim a As Long           'For Our Black/White border
 Dim B As Long           'For Our Black/White border
 Dim col As Byte         'Holds the actual bordercolor
 Dim colOld As Byte      'Also needed for bordercolor
 Dim SC  As Byte         'For Color
 
 
 'Remove Selection if there is one
 PicOrg.Cls
 OrOrNot = False
 TmrShowBorder.Enabled = False

 'How long does it take ?
 Tmr = GetTickCount

 'Get the Picture
 Pic2Array PicOrg, Buf1
 'Create a clear (Black) Picture
 ReDim Buf2(UBound(Buf1, 1), UBound(Buf1, 2))
  
 'We wanted Brightness/Hue instead of RGB
 'Normaly you should write a extra part for brightness/hue
 'and dont do it by only changing to Monochrome/Hue and
 'use the same routine we use for rgb
 'But this is no complete Painter
 'It´s only a demo
 Select Case CboStyle.ListIndex
 Case 1
  PicMonochrome Buf1()
 Case 2
  PicHue Buf1() 'Max 240 colors not 255 like normaly
 End Select

 'Get the Original Pixel Colors
 Ro = Buf1(X, PicOrg.Height - Y).rgbRed
 Go = Buf1(X, PicOrg.Height - Y).rgbGreen
 Bo = Buf1(X, PicOrg.Height - Y).rgbBlue

 'Show the color
 ShpColor.FillColor = PicOrg.Point(X, Y)

 'Allowed RGB Difference
 Percent = HscrPerc.Value

 If CboStyle.ListIndex = 0 Then
 
 'Move thru the Picture and make a blue Point
 'everytime we find a similar color
 For Xl = 0 To UBound(Buf1, 1)
  For Yl = 0 To UBound(Buf1, 2)
   If SimilarColor(Buf1(Xl, Yl).rgbRed, Buf1(Xl, Yl).rgbGreen, Buf1(Xl, Yl).rgbBlue, Ro, Go, Bo, Percent) Then
    Buf2(Xl, Yl).rgbBlue = 255
   End If
  Next Yl
 Next Xl
 
 ElseIf CboStyle.ListIndex < 3 Then
 
 'For Monochrome Pictures i only scann the Red channel
  If CboStyle.ListIndex = 1 Then
   Percent = Percent * 2.55
  Else
   '2.4 because Hue has only 240 colors
   Percent = Percent * 2.4
  End If
  
 For Xl = 0 To UBound(Buf1, 1)
  For Yl = 0 To UBound(Buf1, 2)
     If Abs(CLng(Buf1(Xl, Yl).rgbRed) - CLng(Ro)) <= Percent Then
      Buf2(Xl, Yl).rgbBlue = 255
     End If
  Next Yl
 Next Xl
 
 Else
 'Color
 'Didnt find any better Idea
 Percent = Percent * 2.55 + 1
 SC = SameColor(Ro, Go, Bo)
 For Xl = 0 To UBound(Buf1, 1)
  For Yl = 0 To UBound(Buf1, 2)
   If SameColor(Buf1(Xl, Yl).rgbRed, Buf1(Xl, Yl).rgbGreen, Buf1(Xl, Yl).rgbBlue) = SC Then
    If Abs(CLng(Buf1(Xl, Yl).rgbRed) - CLng(Ro)) < Percent Then
      Buf2(Xl, Yl).rgbBlue = 255
    End If
   End If
  Next Yl
 Next Xl
 
 End If

 'Show the result
 Array2Pic PicDest, Buf2

 'Start a Floodfill on the Original pixels position
 'so we get a white part on the picture
 'Picdest´s fillcolor has to be White
 FloodFill PicDest.HDC, X, Y, vbBlack

 'Make a new black Picture
 ReDim Buf1(UBound(Buf2, 1), UBound(Buf2, 2))
 'Get the floddfilled picture
 Pic2Array PicDest, Buf2




 'Draw the Border
 col = &HFE
 For Xl = 0 To UBound(Buf1, 1)
  a = a + 1
  If a = 5 Then
   a = 0
   col = Not col
  End If
  colOld = col
  For Yl = 0 To UBound(Buf1, 2)
   B = B + 1
   If B = 5 Then
    B = 0
    col = Not col
   End If
   If Buf2(Xl, Yl).rgbRed = 255 Then
    If Xl = 0 Or Xl = UBound(Buf1, 1) Or Yl = 0 Or Yl = UBound(Buf1, 2) Then
     Buf1(Xl, Yl).rgbBlue = col
     Buf1(Xl, Yl).rgbGreen = col
     Buf1(Xl, Yl).rgbRed = col
    ElseIf Buf2(Xl - 1, Yl).rgbRed = 0 Or Buf2(Xl + 1, Yl).rgbRed = 0 Or Buf2(Xl, Yl - 1).rgbRed = 0 Or Buf2(Xl, Yl + 1).rgbRed = 0 Then
     Buf1(Xl, Yl).rgbBlue = col
     Buf1(Xl, Yl).rgbGreen = col
     Buf1(Xl, Yl).rgbRed = col
    End If
   End If
  Next Yl
  B = 0
  col = colOld
 Next Xl

 'Set the Borderpicture
 Array2Pic PicBorder, Buf1

 
'#########################################################
'If you want it as region then remove the '*'
'modRegionShape2.bas is Originaly from LaVolpe
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54017&lngWId=1
'I only changed the some lines
'Old: If transColor = tgtColor No Region for tgtColor
'New: If transColor <> tgtColor Region for tgtColor
'Removed the SetWindowRgn
'Result is now the Region and not True/False

'*' MasterRgn = CreateShapedRegion2(PicDest.Picture.handle, Me.hwnd, &HFFFFFF)
'#########################################################



 'Show how long it took
 LblTime = "Created in " & GetTickCount - Tmr & "ms"

 'Show the Border
 TmrShowBorder.Enabled = True

End Sub

'Load a Picture
Private Sub PicOrg_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Data.Files.Count Then
  On Error Resume Next
  Set PicOrg.Picture = LoadPicture(Data.Files(1))
  If Err Then
   MsgBox "Cant load " & vbCrLf & Data.Files(1), vbCritical + vbOKOnly
   Err.Clear
  Else
   TmrShowBorder.Enabled = False
   X = (PicOrg.Width + 25) * Screen.TwipsPerPixelX
   If X < 490 * Screen.TwipsPerPixelX Then X = 490 * Screen.TwipsPerPixelX
   Me.Width = X
   Me.Height = (PicOrg.Height + 130) * Screen.TwipsPerPixelY
   PicDest.Width = PicOrg.Width
   PicDest.Height = PicOrg.Height
   PicBorder.Width = PicOrg.Width
   PicBorder.Height = PicOrg.Height

  End If
 End If
End Sub

Private Sub TmrShowBorder_Timer()
 Dim Transcol As Long

 'Switch between Black or White as transparent color
 If OrOrNot Then
  Transcol = &HFFFFFF
 End If
 OrOrNot = Not OrOrNot

 'Copy the Border to our Origin
 TransparentBlt PicOrg.HDC, 0, 0, PicOrg.Width, PicOrg.Height, PicBorder.HDC, 0, 0, PicOrg.Width, PicOrg.Height, Transcol
 'Invert the Border so it looks like its moving
 BitBlt PicBorder.HDC, 0, 0, PicBorder.Width, PicBorder.Height, PicBorder.HDC, 0, 0, PATINVERT
End Sub


