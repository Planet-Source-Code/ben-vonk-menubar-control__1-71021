VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Constants
Private Const PIXELS_PER_BITBLT  As Integer = 1

' Private Variables
Private m_CaptionAlignment       As AlignmentConstants
Private m_ItemAlignment          As AlignmentConstants
Private m_IconAlignment          As Alignments
Private m_Animation              As Boolean
Private m_ButtonHideInSingleMenu As Boolean
Private m_FontBoldButtonCaption  As Boolean
Private m_FontBoldItemCaption    As Boolean
Private m_NumberOfMenusChanged   As Boolean
Private m_OnlyFullItemsHit       As Boolean
Private m_OnlyFullItemsShow      As Boolean
Private m_SoundItemScroll        As Boolean
Private m_CurrentMenu            As Integer
Private m_ItemIconSize           As Integer
Private m_ItemsShowed            As Integer
Private PreviousMenu             As Integer
Private m_ButtonHeight           As Long
Private m_ForeColor              As Long
Private m_Menus                  As New Collection
Private picCache                 As PictureBox
Private picMenu                  As PictureBox

' Private API
Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Property Get Animation() As Boolean

   Animation = m_Animation

End Property

Public Property Let Animation(ByVal NewAnimation As Boolean)

   m_Animation = NewAnimation

End Property

Public Property Get ButtonHideInSingleMenu() As Boolean

   ButtonHideInSingleMenu = m_ButtonHideInSingleMenu

End Property

Public Property Let ButtonHideInSingleMenu(ByVal NewButtonHideInSingleMenu As Boolean)

   m_ButtonHideInSingleMenu = NewButtonHideInSingleMenu

End Property

Public Property Get CaptionAlignment() As AlignmentConstants

   CaptionAlignment = m_CaptionAlignment

End Property

Public Property Let CaptionAlignment(ByVal NewCaptionAlignment As AlignmentConstants)

   m_CaptionAlignment = NewCaptionAlignment

End Property

Public Property Let CurrentMenu(ByVal NewCurrentMenu As Long)

   m_CurrentMenu = NewCurrentMenu

End Property

Public Property Get FontBoldButtonCaption() As Boolean

   FontBoldButtonCaption = m_FontBoldButtonCaption

End Property

Public Property Let FontBoldButtonCaption(ByVal NewFontBoldButtonText As Boolean)

   m_FontBoldButtonCaption = NewFontBoldButtonText

End Property

Public Property Get FontBoldItemCaption() As Boolean

   FontBoldItemCaption = m_FontBoldItemCaption

End Property

Public Property Let FontBoldItemCaption(ByVal NewFontBoldItemCaption As Boolean)

   m_FontBoldItemCaption = NewFontBoldItemCaption

End Property

Public Property Let ForeColor(ByVal NewForeColor As Long)

   m_ForeColor = NewForeColor

End Property

Public Property Let ButtonHeight(ByVal NewButtonHeight As Long)

   m_ButtonHeight = NewButtonHeight

End Property

Public Property Set Cache(ByRef NewCache As PictureBox)

   Set picCache = NewCache

End Property

Public Property Get Count() As Long

   Count = m_Menus.Count

End Property

Public Property Get IconAlignment() As Alignments

   IconAlignment = m_IconAlignment

End Property

Public Property Let IconAlignment(ByVal NewIconAlignment As Alignments)

   m_IconAlignment = NewIconAlignment

End Property

Public Property Get Item(ByVal NewIndex As Integer) As clsMenu

   If NewIndex > 0 Then Set Item = m_Menus.Item(NewIndex)

End Property

Public Property Let ItemAlignment(ByVal NewItemAlignment As AlignmentConstants)

   m_ItemAlignment = NewItemAlignment

End Property

Public Property Let ItemIconSize(ByVal NewItemIconSize As Integer)

   m_ItemIconSize = NewItemIconSize

End Property

Public Property Get ItemsShowed() As Integer

   ItemsShowed = m_ItemsShowed

End Property

Public Property Set Menu(ByRef NewMenu As PictureBox)

   Set picMenu = NewMenu

End Property

Public Property Get MenuItems(ByVal SelectedMenu As Integer) As Integer

   MenuItems = m_Menus.Item(SelectedMenu).MenuItemCount

End Property

Public Property Let NumberOfMenusChanged(ByVal NewNumberOfMenusChanged As Boolean)

   m_NumberOfMenusChanged = NewNumberOfMenusChanged

End Property

Public Property Get OnlyFullItemsHit() As Boolean

   OnlyFullItemsHit = m_OnlyFullItemsHit

End Property

Public Property Let OnlyFullItemsHit(ByVal NewOnlyFullItemsHit As Boolean)

   m_OnlyFullItemsHit = NewOnlyFullItemsHit

End Property

Public Property Get OnlyFullItemsShow() As Boolean

   OnlyFullItemsShow = m_OnlyFullItemsShow

End Property

Public Property Let OnlyFullItemsShow(ByVal NewOnlyFullItemsShow As Boolean)

   m_OnlyFullItemsShow = NewOnlyFullItemsShow

End Property

Public Property Get SoundItemScroll() As Boolean

   SoundItemScroll = m_SoundItemScroll

End Property

Public Property Let SoundItemScroll(ByVal NewSoundItemScroll As Boolean)

   m_SoundItemScroll = NewSoundItemScroll

End Property

Public Property Get TopItem() As Integer

   TopItem = m_Menus.Item(m_CurrentMenu).TopMenuItem

End Property

Public Property Get TotalMenuItems() As Integer

Dim clsMenus As clsMenu

   For Each clsMenus In m_Menus
      TotalMenuItems = TotalMenuItems + clsMenus.MenuItemCount
   Next 'clsMenus
   
   Set clsMenus = Nothing

End Property

Public Function Add(ByVal Caption As String, ByVal Index As Integer, ByRef picMenu As PictureBox) As clsMenu

Dim clsNewMenu As New clsMenu

   With clsNewMenu
      .Caption = Caption
      .Index = Index
      Set .Control = picMenu
      .ButtonHeight = m_ButtonHeight
   End With
   
   With m_Menus
      If .Count = 0 Then
         .Add clsNewMenu
         
      ElseIf Index = .Count + 1 Then
         .Add clsNewMenu
         
      ElseIf Index = 1 Then
         .Add clsNewMenu, , 1
         
      Else
         .Add clsNewMenu, , , Index - 1
      End If
   End With
   
   Set Add = clsNewMenu
   Set clsNewMenu = Nothing

End Function

Public Function MouseProcess(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long, Optional ByRef HitType As Integer) As Integer

Const ARROW_DOWN        As Integer = 4
Const ARROW_UP          As Integer = 5
Const SOUND_ITEM_SCROLL As Integer = 2

Dim intResult           As Integer

   intResult = IsMenuButtonHitOrClicked(MousePosition, X, Y)
   
   If intResult Then
      HitType = HIT_TYPE_MENU_BUTTON
      MouseProcess = intResult
      Exit Function
      
   Else
      intResult = m_Menus.Item(m_CurrentMenu).MenuItems.MouseProcess(MousePosition, X, Y)
      
      If intResult Then
         picMenu.ToolTipText = m_Menus.Item(m_CurrentMenu).MenuItems.Item(intResult).ToolTipText
         HitType = HIT_TYPE_MENU_ITEM
         
         If MousePosition = MOUSE_UP Then MouseProcess = intResult
      End If
   End If
   
   intResult = m_Menus.Item(m_CurrentMenu).MouseProcessForArrows(MousePosition, X, Y)
   
   If intResult Then
      If intResult = HIT_TYPE_ARROW Then
         HitType = HIT_TYPE_ARROW
         MouseProcess = HitType
         
      ElseIf intResult = SCROLL_UP Then
         Call ScrollUp
         
         If m_SoundItemScroll Then Call PlaySound(SOUND_ITEM_SCROLL)
         
         HitType = ARROW_UP
         
      ' SCROLL_DOWN
      Else
         Call ScrollDown
         
         If m_SoundItemScroll Then Call PlaySound(SOUND_ITEM_SCROLL)
         
         HitType = ARROW_DOWN
      End If
   End If

End Function

Public Sub Delete(ByVal Index As Integer)

   m_Menus.Remove Index

End Sub

Public Sub MoveToItem(ByVal Moves As Integer)

   If Moves < 0 Then
      For Moves = Abs(Moves) To 1 Step -1
         Call ScrollUp
      Next 'Moves
      
   Else
      For Moves = 1 To Moves
         Call ScrollDown
      Next 'Moves
   End If

End Sub

Public Sub Paint()

   If PreviousMenu = 0 Then PreviousMenu = m_CurrentMenu
   
   If PreviousMenu = m_CurrentMenu Then
      Call RePaint
      
   ElseIf PreviousMenu < m_CurrentMenu Then
      Call ReSelectDown
      
   Else
      Call ReSelectUp
   End If
   
   Call DrawItemIcon
   Call SetButtonsHotSpot
   
   PreviousMenu = m_CurrentMenu

End Sub

Public Sub RestoreButton(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long)

   IsMenuButtonHitOrClicked MousePosition, X, Y

End Sub

Private Function AlignCaption(ByVal CaptionWidth As Long, ByVal ButtonWidth As Long) As Long

   If m_CaptionAlignment = vbLeftJustify Then
      AlignCaption = 5
      
   ElseIf m_CaptionAlignment = vbRightJustify Then
      AlignCaption = ButtonWidth - CaptionWidth - 5
      
   ' vbCenter
   Else
      AlignCaption = (ButtonWidth - CaptionWidth) \ 2
   End If

End Function

Private Function ClipHeight() As Long

   ClipHeight = picMenu.ScaleHeight - (m_Menus.Count - m_CurrentMenu) * m_ButtonHeight

End Function

Private Function FirstIcon() As Long

Dim intMenu As Integer

   For intMenu = 1 To m_CurrentMenu - 1
      FirstIcon = FirstIcon + m_Menus.Item(intMenu).MenuItemCount
   Next 'intMenu
   
   FirstIcon = FirstIcon + m_Menus.Item(m_CurrentMenu).TopMenuItem - 1

End Function

Private Function GetSpace() As Single

   If m_ItemIconSize = 16 Then
      GetSpace = 1.5
      
   Else
      GetSpace = 1
   End If

End Function

Private Function IsMenuButtonHitOrClicked(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long) As Integer

Static intPrevMenu As Integer

Dim clsMenus       As clsMenu
Dim clsPrevMenu    As clsMenu

   If m_ButtonHideInSingleMenu Then Exit Function
   
   For Each clsMenus In m_Menus
      picMenu.ToolTipText = ""
      
      With clsMenus
         If intPrevMenu And (intPrevMenu <> .Index) And (MousePosition <> MOUSE_CHECK) Then
            Call .DrawBorder(RAISED)
            
            intPrevMenu = 0
         End If
         
         If .IsMenuSelected(X, Y) Then
            If MousePosition <> MOUSE_CHECK Then
               If Not clsPrevMenu Is Nothing Then If Y >= picMenu.ScaleHeight \ 2 Then clsPrevMenu.ClearItemHit MousePosition, X, Y
               
               .HitTest MousePosition, X, Y
               picMenu.ToolTipText = .ToolTipText
            End If
            
            IsMenuButtonHitOrClicked = .Index
            intPrevMenu = .Index
            Exit For
         End If
      End With
      
      Set clsPrevMenu = clsMenus
   Next 'clsMenus
   
   Set clsMenus = Nothing
   Set clsPrevMenu = Nothing

End Function

Private Sub DrawItemIcon()

   m_ItemsShowed = m_Menus.Item(m_CurrentMenu).PaintItems(FirstIcon, m_CurrentMenu, ClipHeight, m_Menus.Count, m_ItemIconSize, m_ItemAlignment, m_OnlyFullItemsShow, m_OnlyFullItemsHit, ButtonHideInSingleMenu)

End Sub

Private Sub PaintButton(ByVal CurrentMenu As Integer, ByVal Y As Long, ByVal Width As Long, ByVal IconSize As Integer, ByVal ButtonHeight As Long, ByRef IconTop As Long)

Dim lngX       As Long
Dim strCaption As String

   If m_ButtonHideInSingleMenu Then Exit Sub
   
   With m_Menus.Item(CurrentMenu)
      If Not .Icon Is Nothing Then
         DrawIconEx picMenu.hDC, 4 + ((Width - IconSize - 8) And (m_IconAlignment = [Right Justify])), IconTop, .Icon.Handle, IconSize, IconSize, 0, 0, DI_NORMAL
         lngX = (16 And (m_ButtonHeight = MENU_BUTTON_MIN_HEIGHT))
         
         If m_IconAlignment = [Left Justify] Then
            If m_CaptionAlignment = vbLeftJustify Then lngX = 34 - lngX
            
         ElseIf m_CaptionAlignment = vbRightJustify Then
            lngX = -34 + lngX
         End If
      End If
   End With
   
   With picMenu
      strCaption = m_Menus.Item(CurrentMenu).Caption
      .CurrentX = lngX + AlignCaption(.TextWidth(strCaption), .Width)
      .CurrentY = Y + (m_ButtonHeight - .TextHeight("X")) \ 2
      picMenu.Print strCaption
      IconTop = IconTop + ButtonHeight
   End With

End Sub

Private Sub RePaint()

Dim intIconSize As Integer
Dim intMenu     As Integer
Dim lngIconTop  As Long
Dim lngMax      As Long
Dim lngSrcDC    As Long

   lngMax = m_Menus.Count
   
   With picMenu
      If m_NumberOfMenusChanged Then
         m_NumberOfMenusChanged = False
         .Cls
      End If
      
      lngSrcDC = picCache.hDC
      
      If (lngMax = 0) Or (.hDC = 0) Or (lngSrcDC = 0) Then Exit Sub
      
      .ForeColor = m_ForeColor
      .FontBold = m_FontBoldButtonCaption
      intIconSize = 32 - (16 And (m_ButtonHeight = MENU_BUTTON_MIN_HEIGHT))
      lngIconTop = 3
      
      For intMenu = 1 To m_CurrentMenu
         BitBlt .hDC, 0, (intMenu - 1) * m_ButtonHeight, .ScaleWidth, m_ButtonHeight, lngSrcDC, 0, 0, vbSrcCopy
         
         Call PaintButton(intMenu, CLng(intMenu - 1) * m_ButtonHeight, .ScaleWidth, intIconSize, m_ButtonHeight, lngIconTop)
      Next 'intMenu
      
      lngIconTop = .ScaleHeight - m_ButtonHeight + 3
      
      For intMenu = lngMax To m_CurrentMenu + 1 Step -1
         BitBlt .hDC, 0, .ScaleHeight - (lngMax - intMenu + 1) * m_ButtonHeight, .ScaleWidth, m_ButtonHeight, lngSrcDC, 0, 0, vbSrcCopy
         
         Call PaintButton(intMenu, CLng(.ScaleHeight - (lngMax - intMenu + 1) * m_ButtonHeight), .ScaleWidth, intIconSize, -m_ButtonHeight, lngIconTop)
      Next 'intMenu
      
      .FontBold = m_FontBoldItemCaption
   End With

End Sub

Private Sub ReSelectDown()

Dim blnFirst      As Boolean
Dim intMaxMenus   As Integer
Dim lngHeight     As Long
Dim lngMenuTop    As Long
Dim lngPixelCount As Long
Dim lngSrcDC      As Long
Dim lngStartY     As Long
Dim lngStopY      As Long

   intMaxMenus = m_Menus.Count
   
   With picMenu
      blnFirst = True
      lngSrcDC = picCache.hDC
      lngStartY = .ScaleHeight - (intMaxMenus - m_CurrentMenu) * m_ButtonHeight
      lngStopY = PreviousMenu * m_ButtonHeight
      lngMenuTop = .ScaleHeight - (intMaxMenus - PreviousMenu) * m_ButtonHeight
      
      If (intMaxMenus = 0) Or (.hDC = 0) Or (lngSrcDC = 0) Then Exit Sub
      
      If m_Animation Then
         Do
            BitBlt .hDC, 0, lngStopY, .ScaleWidth, lngStartY - lngStopY - lngPixelCount - PIXELS_PER_BITBLT - ((Not (blnFirst)) * PIXELS_PER_BITBLT), .hDC, 0, lngStopY + PIXELS_PER_BITBLT, vbSrcCopy
            
            If blnFirst Then
               BitBlt .hDC, 0, lngStartY - PIXELS_PER_BITBLT, .ScaleWidth, PIXELS_PER_BITBLT, lngSrcDC, 0, m_ButtonHeight + 3, vbSrcCopy
               blnFirst = False
            End If
            
            lngPixelCount = lngPixelCount + PIXELS_PER_BITBLT
         Loop Until lngMenuTop - ((lngPixelCount + 1) * PIXELS_PER_BITBLT) <= lngStopY
         
         BitBlt .hDC, 0, lngStopY, .ScaleWidth, lngStartY - lngStopY - 1 - lngPixelCount - (Not (blnFirst) * PIXELS_PER_BITBLT), .hDC, 0, lngMenuTop - lngPixelCount, vbSrcCopy
         
      Else
         lngHeight = (m_CurrentMenu - PreviousMenu) * m_ButtonHeight
         BitBlt .hDC, 0, lngStopY, .ScaleWidth, lngHeight, .hDC, 0, lngMenuTop, vbSrcCopy
         StretchBlt .hDC, 0, lngStopY + lngHeight, .ScaleWidth, lngStartY - lngStopY - lngHeight, .hDC, 0, lngMenuTop - 1, 1, 1, vbSrcCopy
      End If
   End With

End Sub

Private Sub ReSelectUp()

Dim blnFirst      As Boolean
Dim intMaxMenus   As Integer
Dim lngMenuBottom As Long
Dim lngPixelCount As Long
Dim lngHeight     As Long
Dim lngSrcDC      As Long
Dim lngStartY     As Long
Dim lngStopY      As Long

   intMaxMenus = m_Menus.Count
   
   With picMenu
      blnFirst = True
      lngSrcDC = picCache.hDC
      lngStartY = m_CurrentMenu * m_ButtonHeight
      lngStopY = .ScaleHeight - (intMaxMenus - PreviousMenu) * m_ButtonHeight
      lngMenuBottom = PreviousMenu * m_ButtonHeight
      
      If (intMaxMenus = 0) Or (.hDC = 0) Or (lngSrcDC = 0) Then Exit Sub
      
      If m_Animation Then
         Do
            BitBlt .hDC, 0, lngStartY + lngPixelCount + PIXELS_PER_BITBLT + ((Not (blnFirst)) * PIXELS_PER_BITBLT), .ScaleWidth, lngStopY - lngStartY - lngPixelCount - PIXELS_PER_BITBLT - ((Not (blnFirst)) * PIXELS_PER_BITBLT), .hDC, 0, lngStartY + lngPixelCount + ((Not (blnFirst)) * PIXELS_PER_BITBLT), vbSrcCopy
            
            If blnFirst Then
               BitBlt .hDC, 0, lngStartY + (lngPixelCount * PIXELS_PER_BITBLT), .ScaleWidth, PIXELS_PER_BITBLT, lngSrcDC, 0, m_ButtonHeight + 3, vbSrcCopy
               blnFirst = False
            End If
            
            lngPixelCount = lngPixelCount + PIXELS_PER_BITBLT
         Loop Until lngMenuBottom + lngPixelCount >= lngStopY
         
         BitBlt .hDC, 0, lngStopY - (PreviousMenu - m_CurrentMenu) * m_ButtonHeight - PIXELS_PER_BITBLT, .ScaleWidth, (PreviousMenu - m_CurrentMenu) * m_ButtonHeight + PIXELS_PER_BITBLT, .hDC, 0, lngStartY + lngPixelCount - PIXELS_PER_BITBLT, vbSrcCopy
         
      Else
         lngHeight = (m_CurrentMenu - PreviousMenu) * m_ButtonHeight
         BitBlt .hDC, 0, lngStopY, .ScaleWidth, lngHeight, .hDC, 0, lngMenuBottom, vbSrcCopy
         StretchBlt .hDC, 0, lngStartY, .ScaleWidth, lngStopY - lngStartY + lngHeight, .hDC, 0, lngMenuBottom + 1, 1, 1, vbSrcCopy
      End If
   End With

End Sub

Private Sub ScrollDown()

Dim intMaxMenus   As Integer
Dim lngStartY     As Long
Dim lngStopY      As Long
Dim lngPixelCount As Long
Dim lngSrcDC      As Long
Dim lngPixel      As Long

   With m_Menus.Item(m_CurrentMenu)
      .TopMenuItem = .TopMenuItem + 1
      .HideButton ARROW_BUTTON_DOWN, m_Menus.Count - m_CurrentMenu
   End With
   
   intMaxMenus = m_Menus.Count
   
   With picMenu
      lngSrcDC = picCache.hDC
      lngStartY = .ScaleHeight - (intMaxMenus - m_CurrentMenu) * m_ButtonHeight
      lngStopY = m_CurrentMenu * m_ButtonHeight
      
      If (intMaxMenus = 0) Or (.hDC = 0) Or (lngSrcDC = 0) Then Exit Sub
      
      If m_Animation Then
         For lngPixel = 1 To m_ItemIconSize * 2 * GetSpace
            ' move the screen up
            BitBlt .hDC, 0, lngStopY, .ScaleWidth, lngStartY - lngStopY, .hDC, 0, lngStopY + PIXELS_PER_BITBLT, vbSrcCopy
            ' repaint the background
            BitBlt .hDC, 0, lngStartY - 1, .ScaleWidth, PIXELS_PER_BITBLT, lngSrcDC, 0, m_ButtonHeight + 3, vbSrcCopy
            lngPixelCount = lngPixelCount + PIXELS_PER_BITBLT
         Next 'lngPixel
      End If
      
      StretchBlt .hDC, 0, lngStopY, .ScaleWidth, lngStartY - lngStopY, .hDC, 0, lngStartY - 1, 1, 1, vbSrcCopy
   End With
   
   Call DrawItemIcon
   Call SetButtonsHotSpot

End Sub

Private Sub ScrollUp()

Dim intMaxMenus   As Integer
Dim lngStartY     As Long
Dim lngStopY      As Long
Dim lngPixelCount As Long
Dim lngSrcDC      As Long
Dim lngPixel      As Long

   With m_Menus.Item(m_CurrentMenu)
      .TopMenuItem = .TopMenuItem - 1
      .HideButton ARROW_BUTTON_UP, m_CurrentMenu
   End With
   
   intMaxMenus = m_Menus.Count
   
   With picMenu
      lngSrcDC = picCache.hDC
      lngStartY = m_CurrentMenu * m_ButtonHeight
      lngStopY = .ScaleHeight - (intMaxMenus - m_CurrentMenu) * m_ButtonHeight
      
      If (intMaxMenus = 0) Or (.hDC = 0) Or (lngSrcDC = 0) Then Exit Sub
      
      If m_Animation Then
         For lngPixel = 1 To m_ItemIconSize * 2 * GetSpace
            ' move the screen up
            BitBlt .hDC, 0, lngStartY + PIXELS_PER_BITBLT, .ScaleWidth, lngStopY - lngStartY - 1, .hDC, 0, lngStartY, vbSrcCopy
            ' repaint the background
            BitBlt .hDC, 0, lngStartY, .ScaleWidth, PIXELS_PER_BITBLT, lngSrcDC, 0, m_ButtonHeight + 3, vbSrcCopy
            lngPixelCount = lngPixelCount + PIXELS_PER_BITBLT
         Next 'lngPixel
      End If
      
      StretchBlt .hDC, 0, lngStopY, .ScaleWidth, lngStartY - lngStopY, .hDC, 0, lngStartY, 1, 1, vbSrcCopy
   End With
   
   Call DrawItemIcon
   Call SetButtonsHotSpot

End Sub

Private Sub SetButtonsHotSpot()

Dim clsMenus As clsMenu
Dim intMax   As Integer

   intMax = m_Menus.Count
   
   For Each clsMenus In m_Menus
      With clsMenus
         If .Index <= m_CurrentMenu Then
            .ButtonTop = (.Index - 1) * m_ButtonHeight
            
         Else
            .ButtonTop = picMenu.ScaleHeight - (intMax - .Index + 1) * m_ButtonHeight
         End If
      End With
   Next 'clsMenus
   
   Set clsMenus = Nothing

End Sub

Private Sub Class_Initialize()

   Set m_Menus = New Collection

End Sub

Private Sub Class_Terminate()

   Set m_Menus = Nothing
   Set picCache = Nothing
   Set picMenu = Nothing

End Sub
