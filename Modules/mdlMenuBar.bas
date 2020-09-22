Attribute VB_Name = "mdlMenuBar"
Option Explicit

' Private Constants
Private Const BF_BOTTOM             As Long = &H8
Private Const BF_LEFT               As Long = &H1
Private Const BF_RIGHT              As Long = &H4
Private Const BF_TOP                As Long = &H2

' Public Constants
Public Const ARROW_BUTTON_DOWN      As Boolean = False
Public Const ARROW_BUTTON_UP        As Boolean = True
Public Const ARROW_BUTTON_SIZE      As Integer = 20
Public Const MENU_BUTTON_MIN_HEIGHT As Integer = 22
Public Const MENU_MAX_ITEMS         As Integer = 15
Public Const MENU_MAX_MENUS         As Integer = 10
Public Const MENU_SPACE             As Integer = 7
Public Const HIT_TYPE_ARROW         As Integer = 3
Public Const HIT_TYPE_MENU_BUTTON   As Integer = 1
Public Const HIT_TYPE_MENU_ITEM     As Integer = 2
Public Const SCROLL_UP              As Integer = 100
Public Const DEFAULT                As Long = 0
Public Const DI_NORMAL              As Long = &H3
Public Const MOUSE_CHECK            As Long = 2
Public Const MOUSE_DOWN             As Long = -1
Public Const MOUSE_MOVE             As Long = 0
Public Const MOUSE_UP               As Long = 1
Public Const RAISED                 As Long = 1
Public Const SUNKEN                 As Long = -1
Public Const BDR_RAISEDINNER        As Long = &H4
Public Const BDR_RAISEDOUTER        As Long = &H1
Public Const BDR_SUNKENINNER        As Long = &H8
Public Const BDR_SUNKENOUTER        As Long = &H2
Public Const BDR_RAISED             As Long = &H5
Public Const BDR_SUNKEN             As Long = &HA
Public Const BF_RECT                As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' Private Enumaration
Private Enum ColorsRGB
   IsRed
   IsGreen
   IsBlue
End Enum

' Public Type
Public Type Rect
   Left                             As Long
   Top                              As Long
   Right                            As Long
   Bottom                           As Long
End Type

' Private Types
Private Type GradientRect
   UpperLeft                        As Long
   LowerRight                       As Long
End Type

Private Type TriVertex
   X                                As Long
   Y                                As Long
   Red                              As Integer
   Green                            As Integer
   Blue                             As Integer
   Alpha                            As Integer
End Type

' Public API's
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal Edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer

' Private API's
Private Declare Function GradientFill Lib "MSImg32" (ByVal hDC As Long, ByRef pVertex As TriVertex, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Private Declare Function OleTranslateColor Lib "OLEPro32" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function sndPlaySound Lib "WinMM" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub DrawGradient(ByRef picRect As Rect, ByRef picObject As Object, ByVal GradientType As GradientButtonTypes, ByVal GradientColor As Long, ByVal BaseColor As Long)

Dim lngRGB         As Long
Dim rctGradient    As GradientRect
Dim tvxGradient(1) As TriVertex

   If GradientType = NoGradient Then Exit Sub
   
   If (GradientType = Right2Left) Or (GradientType = Bottom2Top) Then
      lngRGB = TranslateColor(GradientColor)
      
   Else
      lngRGB = TranslateColor(BaseColor)
   End If
   
   With tvxGradient(0)
      .X = picRect.Left
      .Y = picRect.Top
      .Red = GetColor(TranslateRGB(lngRGB, IsRed))
      .Green = GetColor(TranslateRGB(lngRGB, IsGreen))
      .Blue = GetColor(TranslateRGB(lngRGB, IsBlue))
   End With
   
   If (GradientType = Right2Left) Or (GradientType = Bottom2Top) Then
      lngRGB = TranslateColor(BaseColor)
      
   Else
      lngRGB = TranslateColor(GradientColor)
   End If
   
   With tvxGradient(1)
      .X = picObject.ScaleX(picRect.Right, picObject.ScaleMode, vbPixels)
      .Y = picObject.ScaleY(picRect.Bottom, picObject.ScaleMode, vbPixels)
      .Red = GetColor(TranslateRGB(lngRGB, IsRed))
      .Green = GetColor(TranslateRGB(lngRGB, IsGreen))
      .Blue = GetColor(TranslateRGB(lngRGB, IsBlue))
   End With
   
   rctGradient.UpperLeft = 1
   rctGradient.LowerRight = 0
   GradientFill picObject.hDC, tvxGradient(0), 4, rctGradient, 1, Abs(GradientType >= Top2Bottom)
   Erase tvxGradient

End Sub

Public Sub PlaySound(ByVal Index As Integer)

Const SND_ASYNC     As Long = &H1
Const SND_MEMORY    As Long = &H4
Const SND_NODEFAULT As Long = &H2

Dim strSoundBuffer  As String

   On Local Error Resume Next
   strSoundBuffer = StrConv(LoadResData(Index, "Sounds"), vbUnicode)
   sndPlaySound strSoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
   On Local Error GoTo 0

End Sub

Private Function GetColor(ByVal IsColor As Integer) As Integer

   GetColor = Val("&H" & Hex((IsColor / &HFF&) * &HFFFF&))

End Function

Private Function TranslateColor(ByVal Colors As OLE_COLOR, Optional ByVal Palette As Long) As Long

   If OleTranslateColor(Colors, Palette, TranslateColor) Then TranslateColor = -1

End Function

Private Function TranslateRGB(ByVal ColorVal As Long, ByVal ColorRGB As ColorsRGB) As Long

Dim strHex As String

   strHex = Trim(Hex(ColorVal))
   TranslateRGB = Val("&H" + UCase(Mid(Right("000000", 6 - Len(strHex)) & strHex, 5 - ColorRGB * 2, 2)))

End Function

