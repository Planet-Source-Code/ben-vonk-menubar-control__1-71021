VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Variables
Private m_ArrowUp       As Boolean
Private m_IsDisplayed   As Boolean
Private m_GradientType  As GradientButtonTypes
Private BorderState     As Long
Private m_BackColor     As Long
Private m_Border        As Long
Private m_ButtonHeight  As Long
Private m_ForeColor     As Long
Private m_GradientColor As Long
Private picParent       As PictureBox
Private ImageRectangle  As Rect

Public Property Let ArrowUp(ByVal NewArrowUp As Boolean)
 
   m_ArrowUp = NewArrowUp

End Property

Public Property Let BackColor(ByVal NewBackColor As Long)

   m_BackColor = NewBackColor

End Property

Public Property Let Border(ByVal NewBorder As Long)

   m_Border = NewBorder

End Property

Public Property Let ButtonHeight(ByVal NewButtonHeight As Long)

   m_ButtonHeight = NewButtonHeight

End Property

Public Property Let ForeColor(ByVal NewForeColor As Long)

   m_ForeColor = NewForeColor

End Property

Public Property Let GradientColor(ByVal NewGradientColor As Long)

   m_GradientColor = NewGradientColor

End Property

Public Property Let GradientType(ByVal NewGradientType As GradientButtonTypes)

   m_GradientType = NewGradientType

End Property

Public Property Set Parent(ByVal NewParent As PictureBox)

   Set picParent = NewParent

End Property

Public Function HitTest(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long) As Boolean

   If m_IsDisplayed Then
      If PtInRect(ImageRectangle, X, Y) Then
         HitTest = True
         
         If MousePosition = MOUSE_UP Then
            If BorderState = SUNKEN Then Call DrawBorder(RAISED)
            
         ElseIf MousePosition = MOUSE_DOWN Then
            If BorderState = RAISED Then Call DrawBorder(SUNKEN)
         End If
         
      ElseIf MousePosition = MOUSE_DOWN Then
         If BorderState = SUNKEN Then Call DrawBorder(RAISED)
         
      ' MOUSE_CHECK or MOUSE_MOVE
      ElseIf BorderState = SUNKEN Then
         Call DrawBorder(RAISED)
      End If
   End If

End Function

Public Sub Hide()

   If m_IsDisplayed Then
      With ImageRectangle
         picParent.Line (.Left, .Top)-(.Right, .Bottom), picParent.BackColor, BF
      End With
      
      m_IsDisplayed = False
   End If

End Sub

Public Sub Show(ByVal Alignment As AlignmentConstants, Optional ByVal MenusAtTop As Long, Optional ByVal MenusAtBottom As Long)
 
   If Not picParent Is Nothing Then
      If Not picParent.Visible Then Exit Sub
      
      With ImageRectangle
         If m_ArrowUp Then
            .Top = MENU_SPACE + MenusAtTop * m_ButtonHeight
            
         Else
            .Top = picParent.ScaleHeight - ARROW_BUTTON_SIZE - MENU_SPACE - MenusAtBottom * m_ButtonHeight
         End If
         
         If Alignment = vbRightJustify Then
            .Left = MENU_SPACE
            
         Else
            .Left = picParent.ScaleWidth - MENU_SPACE - ARROW_BUTTON_SIZE
         End If
         
         .Right = .Left + ARROW_BUTTON_SIZE
         .Bottom = .Top + ARROW_BUTTON_SIZE
         
         If .Left Then
            m_IsDisplayed = True
            
            Call DrawBorder(RAISED)
         End If
      End With
   End If

End Sub

Private Sub DrawBorder(ByVal Edge As Long)

Dim lngEdge  As Long
Dim intSize  As Integer
Dim lngLines As Long

   If m_IsDisplayed Then
      With ImageRectangle
         If m_GradientType Then
            Call DrawGradient(ImageRectangle, picParent, m_GradientType, m_GradientColor, m_BackColor)
            
         Else
            picParent.Line (.Left, .Top)-(.Right - 2, .Bottom - 2), m_BackColor, BF
         End If
         
         If Edge = RAISED Then
            If m_Border = BDR_RAISED Then
               lngEdge = BDR_RAISED
               
            Else
               lngEdge = BDR_RAISEDOUTER
            End If
            
            DrawEdge picParent.hDC, ImageRectangle, lngEdge, BF_RECT
            BorderState = RAISED
            
         ' SUNKEN
         Else
            If m_Border = BDR_RAISED Then
               lngEdge = BDR_SUNKEN
               
            Else
               lngEdge = BDR_SUNKENOUTER
            End If
            
            DrawEdge picParent.hDC, ImageRectangle, lngEdge, BF_RECT
            BorderState = SUNKEN
         End If
         
         intSize = 3 - 3 * Abs(m_ArrowUp)
         
         For lngLines = .Top + 8 To .Top + 11
            picParent.Line (.Left - intSize + 9, lngLines)-(.Left + intSize + 11, lngLines), m_ForeColor
            intSize = intSize + Abs(m_ArrowUp) - Abs(Not m_ArrowUp)
         Next 'lnglines
      End With
   End If

End Sub

Private Sub Class_Terminate()

   Set picParent = Nothing

End Sub
