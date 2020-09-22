VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Type
Private Type ButtonStruct
   Rect                  As Rect
   State                 As Long
   OnScreen              As Boolean
End Type

' Private Variables
Private m_ItemAlignement As AlignmentConstants
Private ButtonDownOnMe   As Boolean
Private m_ItemValue      As Boolean
Private m_ButtonStruct   As ButtonStruct
Private m_Index          As Integer
Private m_ItemType       As ItemTypes
Private m_ButtonHeight   As Long
Private m_ItemForeColor  As Long
Private picIcon          As StdPicture
Private picCache         As PictureBox
Private picMenu          As PictureBox
Private HitBorder        As Rect
Private IconBorder       As Rect
Private m_Caption        As String
Private m_Key            As String
Private m_Tag            As String
Private m_ToolTipText    As String

' Private API's
Private Declare Function CreateRectRgnIndirect Lib "GDI32" (lpRect As Rect) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hMF As Long) As Long
Private Declare Function RestoreDC Lib "GDI32" (ByVal hDC As Long, ByVal SavedDC As Long) As Long
Private Declare Function SaveDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function SelectClipRgn Lib "GDI32" (ByVal hDC As Long, ByVal hRgn As Long) As Long

Public Property Let ButtonHeight(ByVal NewButtonHeight As Long)

   m_ButtonHeight = NewButtonHeight

End Property

Public Property Set Cache(ByRef NewCache As PictureBox)

   Set picCache = NewCache

End Property

Public Property Get Caption() As String

   Caption = m_Caption

End Property

Public Property Let Caption(ByVal NewCaption As String)

   m_Caption = NewCaption

End Property

Public Property Get Icon() As StdPicture

   Set Icon = picIcon

End Property

Public Property Set Icon(ByRef NewIcon As StdPicture)

   Set picIcon = NewIcon

End Property

Public Property Get Index() As Integer

   Index = m_Index

End Property

Public Property Let Index(ByVal NewIndex As Integer)

   m_Index = NewIndex

End Property

Public Property Let ItemAlignment(ByVal NewItemAlignment As AlignmentConstants)

   m_ItemAlignement = NewItemAlignment

End Property

Public Property Let ItemForeColor(ByVal NewItemForeColor As Long)

   m_ItemForeColor = NewItemForeColor

End Property

Public Property Get ItemType() As ItemTypes

   ItemType = m_ItemType

End Property

Public Property Let ItemType(ByVal NewItemType As ItemTypes)

   m_ItemType = NewItemType

End Property

Public Property Get ItemValue() As Boolean

   ItemValue = m_ItemValue

End Property

Public Property Let ItemValue(ByVal NewItemValue As Boolean)

   m_ItemValue = NewItemValue
   
   If Not NewItemValue Then Call DrawBorder(DEFAULT)

End Property

Public Property Get Key() As String

   Key = m_Key

End Property

Public Property Let Key(ByVal NewKey As String)

   m_Key = NewKey

End Property

Public Property Set Parent(ByRef NewParent As PictureBox)

   Set picMenu = NewParent

End Property

Public Property Let State(ByVal NewState As Long)

   m_ButtonStruct.State = NewState

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

Public Function DisableButton()

   m_ButtonStruct.OnScreen = False

End Function

Public Function HitTest(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long) As Boolean

   If Not m_ButtonStruct.OnScreen Then Exit Function
   
   If PtInRect(HitBorder, X, Y) Then
      HitTest = True
      
      If MousePosition = MOUSE_UP Then
         If m_ItemValue Then
            Call DrawBorder(SUNKEN)
            
         ElseIf m_ButtonStruct.State <> RAISED Then
            If (m_ItemType <> CheckButton) And (m_ItemType <> OptionButton) Then Call DrawBorder(RAISED)
         End If
         
         ButtonDownOnMe = False
         
      ElseIf MousePosition = MOUSE_DOWN Then
         If m_ButtonStruct.State = DEFAULT Then
            If Not m_ItemValue Then Call DrawBorder(RAISED)
            
         ElseIf m_ButtonStruct.State <> SUNKEN Then
            Call DrawBorder(SUNKEN)
            
            ButtonDownOnMe = True
         End If
         
      ' MOUSE_CHECK or MOUSE_MOVE
      ElseIf Not m_ItemValue Then
         ButtonDownOnMe = False
         
         If m_ButtonStruct.State = DEFAULT Then Call DrawBorder(RAISED)
      End If
      
   Else
      If MousePosition <> MOUSE_MOVE Then ButtonDownOnMe = False
      If Not m_ItemValue And (m_ButtonStruct.State <> DEFAULT) Then Call DrawBorder(DEFAULT)
   End If

End Function

Public Function PaintButton(ByVal TopMenuItem As Integer, ByVal FirstIcon As Long, ByVal CurrentMenu As Integer, ByVal ClipHeight As Long, ByVal ItemIconSize As Integer, ByVal OnlyFullItemsShow As Boolean, ByVal OnlyFullItemsHit As Boolean, ByVal ButtonHideInSingleMenu As Boolean, ByRef IsShowed As Boolean) As Boolean

Dim blnClipping  As Boolean
Dim lngCenter    As Long
Dim lngHeight    As Long
Dim lngLeft      As Long
Dim lngMaxHeight As Long
Dim lngTopItem   As Long
Dim lngRegion    As Long
Dim lngSavedDC   As Long
Dim rctButton    As Rect
Dim rctCaption   As Rect
Dim rctRegion    As Rect
Dim sngSpace     As Single

   PaintButton = True
   IsShowed = False
   
   If m_Index < TopMenuItem Then
      m_ButtonStruct.OnScreen = False
      Exit Function
   End If
   
   With picMenu
      lngTopItem = m_Index - TopMenuItem + 1
      sngSpace = 1 + (MENU_SPACE And (ItemIconSize = 16)) / 10
      rctCaption.Bottom = .TextHeight("X")
      rctCaption.Right = .TextWidth(m_Caption)
      
      If m_ItemAlignement = vbCenter Then
         lngLeft = .ScaleWidth \ 2
         
      ElseIf m_ItemAlignement = vbLeftJustify Then
         lngLeft = ItemIconSize \ 2 + MENU_SPACE
         
      Else  'vbRightJustify
         lngLeft = .ScaleWidth - ItemIconSize \ 2 - MENU_SPACE
      End If
   End With
   
   With m_ButtonStruct
      With .Rect
         If m_ItemAlignement = vbCenter Then
            .Top = m_ButtonHeight - MENU_BUTTON_MIN_HEIGHT + (lngTopItem * 2 * ItemIconSize * sngSpace) - ItemIconSize + ((lngTopItem + 1 = 1) * 4) + (CurrentMenu - 1) * m_ButtonHeight
            
         ' vbLeftJustify or vbRightJustify
         Else
            .Top = m_ButtonHeight - MENU_BUTTON_MIN_HEIGHT \ 3 + (lngTopItem * 1.5 * ItemIconSize * sngSpace) - ItemIconSize + ((lngTopItem + 1 = 1) * 4) + (CurrentMenu - 1) * m_ButtonHeight
         End If
         
         If ButtonHideInSingleMenu Then .Top = .Top - m_ButtonHeight
         
         .Bottom = .Top + ItemIconSize
         .Left = lngLeft - (ItemIconSize \ 2)
         .Right = .Left + ItemIconSize
         rctButton.Left = .Left
         rctButton.Top = .Top
         rctButton.Right = .Right
         rctButton.Bottom = .Bottom
      End With
      
      If rctButton.Top + MENU_SPACE > ClipHeight Then
         .OnScreen = False
         PaintButton = False
         Exit Function
      End If
      
      If rctButton.Bottom > ClipHeight Then
         blnClipping = True
         rctButton.Bottom = ClipHeight
      End If
      
      .OnScreen = True
   End With
   
   If m_ItemAlignement = vbCenter Then
      rctCaption.Left = lngLeft - (rctCaption.Right \ 2)
      rctCaption.Top = rctButton.Top + ItemIconSize + 4
      lngMaxHeight = rctCaption.Top + rctCaption.Bottom
      
   Else
      If m_ItemAlignement = vbLeftJustify Then
         rctCaption.Left = lngLeft + ItemIconSize \ 2 + MENU_SPACE
         
      ' vbRightJustify
      Else
         rctCaption.Left = lngLeft - ItemIconSize \ 2 - rctCaption.Right - MENU_SPACE
      End If
      
      ' vbLeftJustify or vbRightJustify
      rctCaption.Top = rctButton.Top + (rctButton.Bottom - rctButton.Top - rctCaption.Bottom) \ 2
      lngMaxHeight = m_ButtonStruct.Rect.Bottom
   End If
   
   With HitBorder
      .Left = rctButton.Left - 4 - (rctCaption.Right And m_ItemAlignement = vbRightJustify)
      .Top = rctButton.Top - 4
      .Right = rctButton.Right + 4 + (rctCaption.Right And m_ItemAlignement = vbLeftJustify)
      .Bottom = rctButton.Bottom + 4 + (rctCaption.Bottom And (m_ItemAlignement = vbCenter))
      
      If blnClipping Or (.Bottom > ClipHeight) Then .Bottom = (rctButton.Bottom And (Not OnlyFullItemsHit And Not OnlyFullItemsShow))
   End With
   
   With IconBorder
      .Left = rctButton.Left - 4
      .Top = rctButton.Top - 4
      .Right = rctButton.Right + 4
      .Bottom = rctButton.Bottom + (4 And Not blnClipping)
   End With
   
   With m_ButtonStruct.Rect
      If OnlyFullItemsShow And (lngMaxHeight > ClipHeight) Then
         lngHeight = 0
         
      ElseIf blnClipping Then
         lngHeight = rctButton.Bottom - rctButton.Top
         
      Else
         lngHeight = ItemIconSize
      End If
      
      If lngHeight Then
         BitBlt picMenu.hDC, .Left, .Top, ItemIconSize, lngHeight, picCache.hDC, 0, m_ButtonHeight * 2 + (FirstIcon + lngTopItem) * ItemIconSize, vbSrcCopy
         IsShowed = True
      End If
   End With
   
   With picMenu
      .CurrentX = rctCaption.Left
      .CurrentY = rctCaption.Top
      .ForeColor = m_ItemForeColor
      
      If m_ItemAlignement = vbCenter Then
         rctButton.Top = .CurrentY
         
      ' vbLeftJustify or vbRightJustify
      Else
         rctButton.Top = rctButton.Top + (ItemIconSize - rctCaption.Bottom) \ 2
      End If
      
      If lngMaxHeight > ClipHeight Then
         If Not OnlyFullItemsShow Then
            rctRegion.Left = 0
            rctRegion.Top = rctButton.Top
            rctRegion.Right = .Width
            rctRegion.Bottom = ClipHeight
            lngSavedDC = SaveDC(.hDC)
            lngRegion = CreateRectRgnIndirect(rctRegion)
            SelectClipRgn .hDC, lngRegion
            .CurrentY = rctRegion.Top
            picMenu.Print m_Caption
            DeleteObject lngRegion
            RestoreDC .hDC, lngSavedDC
         End If
         
         PaintButton = False
         
      Else
         picMenu.Print m_Caption
         
         If Not blnClipping Then
            PaintButton = True
            
         Else
            PaintButton = False
         End If
      End If
   End With
   
   If ItemValue Then Call DrawBorder(SUNKEN)

End Function

Public Sub DrawBorder(ByVal Edge As Long)

   If Not m_ButtonStruct.OnScreen Then Exit Sub
   
   State = Edge
   
   If Edge = RAISED Then
      DrawEdge picMenu.hDC, IconBorder, BDR_RAISEDOUTER, BF_RECT
      
   ElseIf Edge = SUNKEN Then
      DrawEdge picMenu.hDC, IconBorder, BDR_SUNKENINNER, BF_RECT
      
   ElseIf Not m_ItemValue Then
      With IconBorder
         picMenu.Line (.Left, .Top)-(.Right - 1, .Bottom - 1), picMenu.BackColor, B
      End With
   End If

End Sub

Private Sub Class_Terminate()

   Set picIcon = Nothing
   Set picCache = Nothing
   Set picMenu = Nothing

End Sub
