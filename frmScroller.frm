VERSION 5.00
Begin VB.Form frmScroller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "sinScroller"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLogoMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   2400
      Picture         =   "frmScroller.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   132
      TabIndex        =   7
      Top             =   1410
      Width           =   1980
   End
   Begin VB.PictureBox picSmallScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   1680
      ScaleHeight     =   135
      ScaleWidth      =   585
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   3720
      Picture         =   "frmScroller.frx":01B7
      ScaleHeight     =   5145
      ScaleWidth      =   4800
      TabIndex        =   4
      Top             =   2400
      Width           =   4800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "END"
      Height          =   375
      Left            =   5670
      TabIndex        =   2
      Top             =   2340
      Width           =   1215
   End
   Begin VB.PictureBox picFont 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   -2580
      Picture         =   "frmScroller.frx":136D
      ScaleHeight     =   343
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   1710
      Width           =   4830
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   0
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   0
      Top             =   30
      Width           =   6885
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   450
         Picture         =   "frmScroller.frx":5A85
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   132
         TabIndex        =   6
         Top             =   270
         Width           =   1980
      End
      Begin VB.PictureBox picSmallFont 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3570
         Picture         =   "frmScroller.frx":6111
         ScaleHeight     =   450
         ScaleWidth      =   2400
         TabIndex        =   3
         Top             =   330
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frmScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Type aRefChar
  xPos As Integer
  yPos As Integer
End Type

Private abort As Boolean
Private Const NOTSRCCOPY = &H330008 ' dest = (NOT source)
Private Const NOTSRCERASE = &H1100A6 ' dest = (NOT src) AND (NOT dest)
Private Const BLACKNESS = &H42 ' dest = BLACK
Private Const DSTINVERT = &H550009 ' dest = (NOT dest)
Private Const MERGECOPY = &HC000CA ' dest = (source AND pattern)
Private Const MERGEPAINT = &HBB0226 ' dest = (NOT source) OR dest
Private Const PATCOPY = &HF00021 ' dest = pattern
Private Const PATINVERT = &H5A0049 ' dest = pattern XOR dest
Private Const PATPAINT = &HFB0A09 ' dest = DPSnoo
Private Const SRCAND = &H8800C6 ' dest = source AND dest
Private Const SRCCOPY = &HCC0020 ' dest = source
Private Const SRCERASE = &H440328 ' dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046 ' dest = source XOR dest
Private Const SRCPAINT = &HEE0086 ' dest = source OR dest
Private Const WHITENESS = &HFF0062  ' dest = WHITE

Private scrollTextBig As String
Private scrollTextSmall As String
Private refCharBig(255) As aRefChar
Private refCharSmall(255) As aRefChar
Private refSin(3600) As Integer
Private firstTime As Boolean

Private Sub Command1_Click()
  abort = True
End Sub

Private Sub Form_Load()
  scrollTextBig = "       I'M BACK WITH SOME FINE SCROLL ROUTINES!!    HOW YA LIKE ME NOW!      BACK IN THE OLD DAYS I USED TO WRITE DEMO'S FOR THE ATARI-ST.  I DIDN'T USE ASSEMBLER BUT JUST PLAIN OMIKRON BASIC."
  scrollTextBig = scrollTextBig & " THIS TIME I USED VISUAL BASIC. VB IS NOT SUPPOSED TO BE USED FOR DEMO PROGRAMMING BUT SOME OLD SKOOL EFFECTS DO WORK GREAT.  "
  scrollTextBig = scrollTextBig & " GREETINGS TO ALL FORMER ATARI ST DEMO PROGRAMMERS AND SEE YOU SOON STNICCC .....   WRAP  .....                "
  scrollTextSmall = "THIS SMALL DEMO WAS MADE BY THE ONE AND ONLY FLYGUY OF THE DOUBLE DUTCH CREW!        "
  initChar
  firstTime = True
  picFont.Visible = False
  picSmallFont.Visible = False
  picMask.Visible = False
  picSmallScroll.Visible = False
  picLogo.Visible = False
  picLogoMask.Visible = False
End Sub

Private Sub Form_Resize()
  picScroll.Move 0, 0, Me.ScaleWidth, 150
  picSmallScroll.Width = Me.ScaleWidth
  picSmallScroll.Height = 12
  picSmallScroll.Left = 0
  firstTime = False
  Scroll
End Sub

Private Sub Scroll()
  Dim nofChar As Integer
  Dim bigFontWidth As Long, bigFontHeight As Long
  Dim bigCharX As Long
  Dim ibigScroll As Long, iViewPort As Integer, iSmallScroll As Integer
  Dim smallCharX As Long
  Dim charPerPage As Long
  Dim retVal As Long
  Dim curChar As Byte
  Dim Y As Long, i As Long, Y2 As Long, X As Long
  Dim deg As Long
  Dim change As Integer
  Dim psScaleWidth As Long, psHDC As Long
  Dim pfHDC As Long, pmHDC As Long, pssHDC As Long
  Dim logoDeg As Long
  Dim yLogo As Long
  Dim xLogo As Long
  Dim xLogoMax As Long, yLogoMax As Long
  Dim yLogoSwap As Long, xLogoSwap As Long
  
  
  bigFontWidth = 64
  bigFontHeight = 49
  picScroll.ScaleMode = vbPixels
  picFont.ScaleMode = vbPixels
  picSmallFont.ScaleMode = vbPixels
  picSmallScroll.ScaleMode = vbPixels
  
  psScaleWidth = picScroll.ScaleWidth
  psHDC = picScroll.hDC
  pfHDC = picFont.hDC
  pmHDC = picMask.hDC
  
  pssHDC = picSmallScroll.hDC
  xLogoMax = psScaleWidth - picLogo.ScaleWidth
  yLogoMax = picScroll.ScaleHeight - picLogo.ScaleHeight
  xLogoSwap = 1
  yLogoSwap = 1
  xLogo = 0
  yLogo = 0
  charPerPage = picScroll.Width \ bigFontWidth + 1
  
  nofChar = Len(scrollTextBig)
  bigCharX = 0
  ibigScroll = 1
  iSmallScroll = 1
  change = 1
  While True
    ' The small scroller
    ' scroll to the left
    picScroll.ClipControls = False
    curChar = Asc(Mid$(scrollTextSmall, iSmallScroll, 1))
    retVal = BitBlt(pssHDC, 0, 0, picSmallScroll.ScaleWidth, 12, _
                    pssHDC, 1, 0, SRCCOPY)
    ' add new piece of font
    retVal = BitBlt(pssHDC, picSmallScroll.ScaleWidth - 1, 0, 1, 12, _
                    picSmallFont.hDC, refCharSmall(curChar).xPos + smallCharX, refCharSmall(curChar).yPos, SRCCOPY)
    ' copy to the screen
    retVal = BitBlt(psHDC, 0, 0, psScaleWidth, 12, pssHDC, 0, 0, SRCCOPY)
    retVal = BitBlt(psHDC, 0, 10, psScaleWidth, 12, psHDC, 0, 0, SRCCOPY)
    retVal = BitBlt(psHDC, 0, 20, psScaleWidth, 24, psHDC, 0, 0, SRCCOPY)
    retVal = BitBlt(psHDC, 0, 40, psScaleWidth, 48, psHDC, 0, 0, SRCCOPY)
    retVal = BitBlt(psHDC, 0, 80, psScaleWidth, 70, psHDC, 0, 0, SRCCOPY)
    smallCharX = smallCharX + 1
    If smallCharX = 15 Then
      smallCharX = 0
      iSmallScroll = iSmallScroll + 1
      If iSmallScroll > Len(scrollTextSmall) Then iSmallScroll = 1
    End If
    
    ' Dancing DDC logo
    
    retVal = BitBlt(psHDC, xLogo, yLogo + i, 132, 64, picLogoMask.hDC, 0, 0, SRCAND)
    retVal = BitBlt(psHDC, xLogo, yLogo + i, 132, 64, picLogo.hDC, 0, 0, SRCPAINT)
    xLogo = xLogo + xLogoSwap
    yLogo = yLogo + yLogoSwap
    If xLogo = 0 Or xLogo = xLogoMax Then xLogoSwap = xLogoSwap * -1
    If yLogo = 0 Or yLogo = yLogoMax Then yLogoSwap = yLogoSwap * -1
    
    ' The BIG scoller
    deg = deg + 10
    If deg > 3600 Then deg = 0
    change = 1
    For iViewPort = 0 To charPerPage
      Y = deg + (ibigScroll + iViewPort) * 200
      If Y > 3600 Then Y = Y Mod 3600
      Y2 = Y
      'y2 = deg + (ibigScroll + iViewPort) * 250
      'If y2 > 3600 Then y2 = y2 Mod 3600
      X = iViewPort * bigFontWidth - bigCharX
      curChar = Asc(Mid$(scrollTextBig, ibigScroll + iViewPort, 1))
      If (ibigScroll + iViewPort) / 2 = (ibigScroll + iViewPort) \ 2 Then
        retVal = BitBlt(psHDC, X, 50 - refSin(Y), bigFontWidth, bigFontHeight, _
                        pmHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCAND)
        retVal = BitBlt(psHDC, X, 50 - refSin(Y), bigFontWidth, bigFontHeight, _
                        pfHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCPAINT)
        retVal = BitBlt(psHDC, X, 50 + refSin(Y2), bigFontWidth, bigFontHeight, _
                        pmHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCAND)
        retVal = BitBlt(psHDC, X, 50 + refSin(Y2), bigFontWidth, bigFontHeight, _
                        pfHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCPAINT)
      Else
        retVal = BitBlt(psHDC, X, 50 + refSin(Y2), bigFontWidth, bigFontHeight, _
                        pmHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCAND)
        retVal = BitBlt(psHDC, X, 50 + refSin(Y2), bigFontWidth, bigFontHeight, _
                        pfHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCPAINT)
        retVal = BitBlt(psHDC, X, 50 - refSin(Y), bigFontWidth, bigFontHeight, _
                        pmHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCAND)
        retVal = BitBlt(psHDC, X, 50 - refSin(Y), bigFontWidth, bigFontHeight, _
                        pfHDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCPAINT)
      End If
    Next iViewPort
    bigCharX = bigCharX + 2
    If bigCharX >= bigFontWidth Then bigCharX = 0
    If bigCharX = 0 Then
      ibigScroll = ibigScroll + 1
      If ibigScroll > nofChar - charPerPage Then ibigScroll = 1
    End If
    '
    picScroll.ClipControls = True
    picScroll.Refresh
    DoEvents
    If abort Then GoTo theEnd
  Wend
theEnd:
  End
End Sub

Private Sub initChar()
  Dim fontChar As String
  Dim i As Integer
  Dim curChar As Byte
  Dim pi As Double
  Dim j As Double
  
  fontChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZ,'()!?.- "
  
  For i = 0 To Len(fontChar) - 1
    curChar = Asc(Mid$(fontChar, i + 1, 1))
    refCharBig(curChar).yPos = Int(i / 5) * 49
    refCharBig(curChar).xPos = (i Mod 5) * 64
  Next i
  
  fontChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZ .!?"
  
  For i = 0 To Len(fontChar) - 1
    curChar = Asc(Mid$(fontChar, i + 1, 1))
    refCharSmall(curChar).yPos = Int(i / 10) * 10
    refCharSmall(curChar).xPos = (i Mod 10) * 16
  Next i
  
  ' use a lookup table in tenth's of degrees
  ' with amplitude already calculated
  For i = 0 To 3600
    refSin(i) = 50 * Sin((i / 3600) * (2 * 3.141592653))
  Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
  End
End Sub

Public Sub OK()
#If NEEDED Then
      retVal = BitBlt(picScroll.hDC, iViewPort * bigFontWidth - bigCharX, 50 - refSin(Y), bigFontWidth, bigFontHeight, _
                      picMask.hDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCAND)
      retVal = BitBlt(picScroll.hDC, iViewPort * bigFontWidth - bigCharX, 50 - refSin(Y), bigFontWidth, bigFontHeight, _
                      picFont.hDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCPAINT)
      retVal = BitBlt(picScroll.hDC, iViewPort * bigFontWidth - bigCharX, 50 - refSin(Y2), bigFontWidth, bigFontHeight, _
                      picMask.hDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCAND)
      retVal = BitBlt(picScroll.hDC, iViewPort * bigFontWidth - bigCharX, 50 - refSin(Y2), bigFontWidth, bigFontHeight, _
                      picFont.hDC, refCharBig(curChar).xPos, refCharBig(curChar).yPos, SRCPAINT)
#End If
End Sub

