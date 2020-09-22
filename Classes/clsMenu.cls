VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Variables
Private m_IconAlignment    As Alignments
Private m_Password         As Boolean
Private m_ArrowDown        As clsArrow
Private m_ArrowUp          As clsArrow
Private m_MenuItems        As clsMenuItems
Private m_Index            As Integer
Private m_TopMenuItem      As Integer
Private m_ButtonHeight     As Long
Private m_Border           As Long
Private picMenu            As PictureBox
Private picCache           As PictureBox
Private m_HotSpot          As Rect
Private m_Icon             As StdPicture
Private m_Caption          As String
Private m_Tag              As String
Private m_ToolTipText      As String

Public Property Let BackColor(ByVal NewBackColor As Long)

   m_ArrowUp.BackColor = NewBackColor
   m_ArrowDown.BackColor = NewBackColor

End Property

Public Property Let Border(ByVal NewBorder As Long)

   m_Border = NewBorder
   m_ArrowUp.Border = NewBorder
   m_ArrowDown.Border = NewBorder

End Property

Public Property Get ButtonHeight() As Long

   ButtonHeight = m_ButtonHeight

End Property

Public Property Let ButtonHeight(ByVal NewButtonHeight As Long)

Dim intItem As Integer

   m_ButtonHeight = NewButtonHeight
   m_ArrowUp.ButtonHeight = NewButtonHeight
   m_ArrowDown.ButtonHeight = NewButtonHeight
   
   With m_MenuItems
      For intItem = 1 To .Count
         .Item(intItem).ButtonHeight = NewButtonHeight
      Next 'intItem
   End With

End Property

Public Property Get ButtonTop() As Long

   ButtonTop = m_HotSpot.Top

End Property

Public Property Let ButtonTop(ByVal NewButtonTop As Long)

   With m_HotSpot
      .Left = 0
      .Top = NewButtonTop
      .Right = picMenu.ScaleWidth
      .Bottom = NewButtonTop + m_ButtonHeight
   End With

End Property

Public Property Get Caption() As String

   Caption = m_Caption

End Property

Public Property Let Caption(ByVal NewCaption As String)

   m_Caption = NewCaption

End Property

Public Property Set Control(ByRef NewControl As PictureBox)

   Set picMenu = NewControl
   Set m_ArrowUp.Parent = NewControl
   Set m_ArrowDown.Parent = NewControl

End Property

Public Property Let ForeColor(ByVal NewForeColor As Long)

   m_ArrowUp.ForeColor = NewForeColor
   m_ArrowDown.ForeColor = NewForeColor

End Property

Public Property Let GradientColor(ByVal NewGradientColor As Long)

   m_ArrowUp.GradientColor = NewGradientColor
   m_ArrowDown.GradientColor = NewGradientColor

End Property

Public Property Let GradientType(ByVal NewGradientType As Long)

   m_ArrowUp.GradientType = NewGradientType
   m_ArrowDown.GradientType = NewGradientType

End Property

Public Property Get Icon() As StdPicture

   Set Icon = m_Icon

End Property

Public Property Set Icon(ByRef NewIcon As StdPicture)

   Set m_Icon = NewIcon

End Property

Public Property Get IconAlignment() As Alignments

   IconAlignment = m_IconAlignment

End Property

Public Property Let IconAlignment(ByRef NewIconAlignment As Alignments)

   m_IconAlignment = NewIconAlignment

End Property

Public Property Set ImageCache(ByRef NewImageCache As PictureBox)

   Set picCache = NewImageCache

End Property

Public Property Get Index() As Integer

   Index = m_Index

End Property

Public Property Let Index(ByVal NewIndex As Integer)

   m_Index = NewIndex

End Property

Public Property Get MenuItems() As clsMenuItems

   Set MenuItems = m_MenuItems

End Property

Public Property Get Password() As Boolean

   Password = m_Password

End Property

Public Property Let Password(ByVal NewPassword As Boolean)

   m_Password = NewPassword

End Property

Public Property Get Tag() As String

   Tag = m_Tag

End Property

Public Property Let Tag(ByVal NewTag As String)

   m_Tag = NewTag

End Property

Public Property Get ToolTipText() As String

   ToolTipText = m_ToolTipText

End Property

Public Property Let ToolTipText(ByVal NewToolTipText As String)

   m_ToolTipText = NewToolTipText

End Property

Public Property Get TopMenuItem() As Integer

   If m_TopMenuItem = 0 Then m_TopMenuItem = 1
   
   TopMenuItem = m_TopMenuItem

End Property

Public Property Let TopMenuItem(ByVal NewTopMenuItem As Integer)

   If NewTopMenuItem Then m_TopMenuItem = NewTopMenuItem

End Property

Public Function AddMenuItem(ByVal Caption As String, ByVal MenuItemlIndex As Integer, ByRef picIcon As StdPicture) As clsMenuItem

   With m_MenuItems
      .Add Caption, MenuItemlIndex, m_ButtonHeight, picIcon
      Set .Item(MenuItemlIndex).Parent = picMenu
      Set .Item(MenuItemlIndex).Cache = picCache
   End With

End Function

Public Function ClearItemHit(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long) As Boolean

Dim intItem As Integer

   With m_MenuItems
      For intItem = 1 To .Count
         .Item(intItem).HitTest MousePosition, X, Y
      Next 'intItem
   End With

End Function

Public Function HitTest(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long) As Boolean

Dim lngEdge As Long

   If MousePosition = MOUSE_DOWN Then
      lngEdge = SUNKEN
      
   Else
      lngEdge = RAISED
   End If
   
   If PtInRect(m_HotSpot, X, Y) Then Call DrawBorder(lngEdge)

End Function

Public Function IsMenuSelected(ByVal X As Long, ByVal Y As Long) As Boolean

   IsMenuSelected = Not (PtInRect(m_HotSpot, X, Y) = 0)
   
   If Err.Number Then
      IsMenuSelected = False
      Err.Clear
   End If

End Function

Public Function MenuItem(ByVal Index As Integer) As clsMenuItem

   Set MenuItem = m_MenuItems.Item(Index)

End Function

Public Function MenuItemCount() As Integer

   MenuItemCount = m_MenuItems.Count

End Function

Public Function MouseProcessForArrows(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long) As Integer

Const BUTTON_DOWN         As Integer = -1
Const BUTTON_UP           As Integer = 1
Const SCROLL_DOWN         As Integer = -100

Static lngPrevPosition(1) As Long

Dim clsUpDownArrow        As clsArrow
Dim intCount              As Integer
Dim lngDirection          As Long

   For intCount = 0 To 1
      If intCount Then
         Set clsUpDownArrow = m_ArrowUp
         
      Else
         Set clsUpDownArrow = m_ArrowDown
      End If
      
      If clsUpDownArrow.HitTest(MousePosition, X, Y) Then
         If MousePosition = MOUSE_UP Then
            If lngPrevPosition(intCount) = BUTTON_DOWN Then
               If intCount Then
                  lngDirection = SCROLL_UP
                  
               Else
                  lngDirection = SCROLL_DOWN
               End If
               
               MouseProcessForArrows = CInt(lngDirection)
            End If
            
            lngPrevPosition(intCount) = MousePosition
            
         ElseIf MousePosition = MOUSE_DOWN Then
            MouseProcessForArrows = HIT_TYPE_ARROW
            lngPrevPosition(intCount) = MousePosition
            
         ' MOUSE_CHECK or MOUSE_MOVE
         Else
            MouseProcessForArrows = HIT_TYPE_ARROW
            
            If lngPrevPosition(intCount) <> BUTTON_DOWN Then lngPrevPosition(intCount) = MousePosition
         End If
         
      ElseIf MousePosition = MOUSE_UP Then
         lngPrevPosition(intCount) = BUTTON_UP
      End If
   Next 'intCount
   
   Set clsUpDownArrow = Nothing

End Function

Public Function PaintItems(ByVal FirstIcon As Long, ByVal CurrentMenu As Integer, ByVal ClipHeight As Long, ByVal Max As Integer, ByVal ItemIconSize As Integer, ByVal ItemIconAlignment As AlignmentConstants, ByVal OnlyFullItemsShow As Boolean, ByVal OnlyFullItemsHit As Boolean, ByVal ButtonHideInSingleMenu As Boolean) As Integer

   m_MenuItems.ItemsShowed = 0
   
   If m_MenuItems.Paint(m_TopMenuItem, FirstIcon, CurrentMenu, ClipHeight, ItemIconSize, OnlyFullItemsShow, OnlyFullItemsHit, ButtonHideInSingleMenu) Then
      m_ArrowDown.Hide
      
   Else
      m_ArrowDown.Show ItemIconAlignment, MenusAtBottom:=Max - CurrentMenu
   End If
   
   If m_TopMenuItem > 1 Then
      m_ArrowUp.Show ItemIconAlignment, MenusAtTop:=CurrentMenu - (CurrentMenu And ButtonHideInSingleMenu)
      
   Else
      m_ArrowUp.Hide
   End If
   
   PaintItems = m_MenuItems.ItemsShowed

End Function

Public Sub DeleteMenuItem(ByVal MenuItemlIndex As Long)

   Call m_MenuItems.Delete(MenuItemlIndex)

End Sub

Public Sub HideButton(ByVal ButtonUp As Boolean, ByVal Offset As Long)

   If ButtonUp Then
      m_ArrowUp.Hide
      
   Else
      m_ArrowDown.Hide
   End If

End Sub

Public Sub DrawBorder(ByVal Edge As Long)

Dim lngEdge As Long

   If Edge = RAISED Then
      lngEdge = m_Border
      
   ' SUNKEN
   ElseIf m_Border = BDR_RAISED Then
      lngEdge = BDR_SUNKEN
      
   Else
      lngEdge = BDR_SUNKENINNER
   End If
   
   DrawEdge picMenu.hDC, m_HotSpot, lngEdge, BF_RECT
 
End Sub

Private Sub Class_Initialize()

   Set m_MenuItems = New clsMenuItems
   Set m_ArrowDown = New clsArrow
   Set m_ArrowUp = New clsArrow
   m_ArrowUp.ArrowUp = ARROW_BUTTON_UP
   m_ArrowDown.ArrowUp = ARROW_BUTTON_DOWN
   m_TopMenuItem = 1

End Sub

Private Sub Class_Terminate()

   Set m_Icon = Nothing
   Set m_MenuItems = Nothing
   Set m_ArrowDown = Nothing
   Set m_ArrowUp = Nothing
   Set picCache = Nothing
   Set picMenu = Nothing

End Sub

