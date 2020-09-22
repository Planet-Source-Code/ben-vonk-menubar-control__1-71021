VERSION 5.00
Begin VB.UserControl MenuBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   624
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1812
   Picture         =   "MenuBar.ctx":0000
   PropertyPages   =   "MenuBar.ctx":030A
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   ToolboxBitmap   =   "MenuBar.ctx":034D
   Begin VB.Timer tmrMouseOut 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   120
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   120
      Width           =   384
   End
   Begin VB.PictureBox picCache 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   960
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "MenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'MenuBar Control
'
'Author Ben Vonk
'03-07-2005 First version
'20-01-2008 Second version, some bugfixes and updated with more features

Option Explicit

' Public Events
Public Event BeforeOpenMenu(MenuIndex As Integer, AlReadyOpen As Boolean, Cancel As Boolean)
Public Event CheckMenuWithPassword(MenuIndex As Integer, Cancel As Boolean)
Public Event LockMenuWithPassword(MenuIndex As Integer, Cancel As Boolean)
Public Event MenuClick(MenuIndex As Integer)
Public Event MenuItemClick(MenuIndex As Integer, ItemIndex As Integer, ItemKey As String, ItemType As ItemTypes, ItemValue As Boolean, PreviousOption As Integer)

' Private Constants
Private Const MENU_BUTTON_CAPTION As String = "Menu"
Private Const MENU_ITEM_CAPTION   As String = "Item"

' Public Enumerations
Public Enum Alignments
   [Left Justify]
   [Right Justify]
End Enum

Public Enum BorderStyles
   None
   [Fixed Single]
   Edged
End Enum

Public Enum ButtonHeights
   Low
   High
End Enum

Public Enum ItemIconSizeType
   [16x16]
   [32x32]
   [48x48]
   [64x64]
End Enum

Public Enum ItemTypes
   DefaultButton
   CheckButton
   OptionButton
   ResetButton
   LockMenuButton
End Enum

Public Enum MenuStyles
   Flat
   [3D]
End Enum

Public Enum GradientButtonTypes
   NoGradient
   Left2Right
   Right2Left
   Top2Bottom
   Bottom2Top
End Enum

' Private Enumeration
Private Enum MenuObjects
   MenuBorder
   MenuButtonBackColor
   MenuButtonForeColor
   MenuButtonGradientColor
   MenuButtonGradientType
   MenuButtonHeight
   MenuItemAlignment
   MenuItemForeColor
End Enum

' Private Type
Private Type PointAPI
   X                              As Long
   Y                              As Long
End Type

' Private Variables
Private m_ItemAlignment           As AlignmentConstants
Private Initializing              As Boolean
Private m_Locked                  As Boolean
Private m_SoundMenuOpen           As Boolean
Private m_ButtonHeight            As ButtonHeights
Private m_Menus                   As clsMenus
Private m_ItemIconSize            As ItemIconSizeType
Private SizeIcon                  As Integer
Private m_CurrentMenu             As Integer
Private m_CurrentItem             As Integer
Private m_MaxMenus                As Integer
Private m_MaxItems                As Integer
Private m_StartupMenu             As Integer
Private MenuHitObject             As Integer
Private MenuHitType               As Integer
Private ToCurrentItem             As Boolean
Private m_ButtonGradientType      As GradientButtonTypes
Private HeightMenuButton          As Long
Private hWndMenuBar               As Long
Private m_ButtonBackColor         As Long
Private m_ButtonForeColor         As Long
Private m_ButtonGradientColor     As Long
Private m_MenuBorder              As Long
Private m_ItemForeColor           As Long
Private MenuObjectX               As Long
Private MenuObjectY               As Long
Private m_MenuButtonIcon          As StdPicture
Private m_ItemIcon                As StdPicture
Private m_BorderStyle             As BorderStyles
Private m_Appearance              As MenuStyles

' Private API's
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Property Get Animation() As Boolean
Attribute Animation.VB_Description = "Determines whether the menu open or item scroll is showed with scrolling animation."

   Animation = m_Menus.Animation

End Property

Public Property Let Animation(ByVal NewAnimation As Boolean)

   m_Menus.Animation = NewAnimation
   PropertyChanged "Animation"

End Property

Public Property Get Appearance() As MenuStyles
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."

   Appearance = m_Appearance

End Property

Public Property Let Appearance(ByVal NewAppearance As MenuStyles)

   m_Appearance = NewAppearance
   
   If m_Appearance Then
      m_MenuBorder = BDR_RAISEDINNER
      
   Else
      m_MenuBorder = BDR_RAISED
   End If
   
   PropertyChanged "Appearance"
   
   Call SetMenuObjectValue(m_MenuBorder, MenuBorder)
   Call SetupCache

End Property

Public Property Get BorderStyle() As BorderStyles
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."

   BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As BorderStyles)

   m_BorderStyle = NewBorderStyle
   
   If (UserControl.BorderStyle = None) And ((m_BorderStyle = None) Or (m_BorderStyle = Edged)) Then Call UserControl_Resize
   
   If m_BorderStyle < Edged Then
      UserControl.BorderStyle = m_BorderStyle
      
   Else
      UserControl.BorderStyle = None
   End If
   
   PropertyChanged "BorderStyle"

End Property

Public Property Get ButtonCaption() As String
Attribute ButtonCaption.VB_Description = "Returns/sets the text displayed in an menu button."

   ButtonCaption = m_Menus.Item(m_CurrentMenu).Caption

End Property

Public Property Let ButtonCaption(ByVal NewButtonCaption As String)

   m_Menus.Item(m_CurrentMenu).Caption = NewButtonCaption
   PropertyChanged "ButtonCaption"
   
   Call SetupCache

End Property

Public Property Get ButtonCaptionAlignment() As AlignmentConstants
Attribute ButtonCaptionAlignment.VB_Description = "Returns/sets the caption alignment of an menu button."

   ButtonCaptionAlignment = m_Menus.CaptionAlignment

End Property

Public Property Let ButtonCaptionAlignment(ByVal NewButtonCaptionAlignment As AlignmentConstants)

   m_Menus.CaptionAlignment = NewButtonCaptionAlignment
   PropertyChanged "ButtonCaptionAlignment"
   
   Call SetupCache

End Property

Public Property Get ButtonBackColor() As OLE_COLOR
Attribute ButtonBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an menu button."

   ButtonBackColor = m_ButtonBackColor

End Property

Public Property Let ButtonBackColor(ByVal NewButtonBackColor As OLE_COLOR)

   m_ButtonBackColor = NewButtonBackColor
   PropertyChanged "ButtonBackColor"
   
   Call SetMenuObjectValue(m_ButtonBackColor, MenuButtonBackColor)
   Call SetupCache

End Property

Public Property Get ButtonForeColor() As OLE_COLOR
Attribute ButtonForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an menu button."

   ButtonForeColor = m_ButtonForeColor

End Property

Public Property Let ButtonForeColor(ByVal NewButtonForeColor As OLE_COLOR)

   m_ButtonForeColor = NewButtonForeColor
   m_Menus.ForeColor = m_ButtonForeColor
   picMenu.Cls
   
   Call SetMenuObjectValue(m_ButtonForeColor, MenuButtonForeColor)
   Call picMenu_Paint
   
   PropertyChanged "ButtonForeColor"

End Property

Public Property Get ButtonGradientColor() As OLE_COLOR
Attribute ButtonGradientColor.VB_Description = "Returns/sets the color used to display the gradient of an menu button."

   ButtonGradientColor = m_ButtonGradientColor

End Property

Public Property Let ButtonGradientColor(ByVal NewButtonGradientColor As OLE_COLOR)

   m_ButtonGradientColor = NewButtonGradientColor
   PropertyChanged "ButtonGradientColor"
   
   Call SetMenuObjectValue(m_ButtonGradientColor, MenuButtonGradientColor)
   Call SetupCache

End Property

Public Property Get ButtonGradientType() As GradientButtonTypes
Attribute ButtonGradientType.VB_Description = "Returns/sets the view type used to display the gradient of an menu button."

   ButtonGradientType = m_ButtonGradientType

End Property

Public Property Let ButtonGradientType(ByVal NewButtonGradientType As GradientButtonTypes)

   m_ButtonGradientType = NewButtonGradientType
   PropertyChanged "ButtonGradientType"
   
   Call SetMenuObjectValue(m_ButtonGradientType, MenuButtonGradientType)
   Call SetupCache

End Property

Public Property Get ButtonHeight() As ButtonHeights
Attribute ButtonHeight.VB_Description = "Returns/sets the height of an menu button."

   ButtonHeight = m_ButtonHeight

End Property

Public Property Let ButtonHeight(ByVal NewButtonHeight As ButtonHeights)

   If NewButtonHeight < Low Then NewButtonHeight = Low
   If NewButtonHeight > High Then NewButtonHeight = High
   
   m_ButtonHeight = NewButtonHeight
   HeightMenuButton = GetMenuButtonHeight(m_ButtonHeight)
   m_Menus.ButtonHeight = HeightMenuButton
   picCache.Height = GetMenuButtonHeight(-1)
   PropertyChanged "ButtonHeight"
   
   Call UserControl_Resize
   Call SetMenuObjectValue(HeightMenuButton, MenuButtonHeight)
   
   With picMenu
      .Cls
      
      If .Font.Size > HeightMenuButton - 10 Then .Font.Size = HeightMenuButton - 10
   End With
   
   Call SetupCache

End Property

Public Property Get ButtonHideInSingleMenu() As Boolean
Attribute ButtonHideInSingleMenu.VB_Description = "Determines whether an button is showed in a single menu."

   ButtonHideInSingleMenu = m_Menus.ButtonHideInSingleMenu

End Property

Public Property Let ButtonHideInSingleMenu(ByVal NewButtonHideInSingleMenu As Boolean)

   If m_MaxMenus > 1 Then NewButtonHideInSingleMenu = False
   
   m_Menus.ButtonHideInSingleMenu = NewButtonHideInSingleMenu
   PropertyChanged "ButtonHideInSingleMenu"
   
   Call UserControl_Resize
   Call SetupCache

End Property

Public Property Get ButtonIcon() As StdPicture
Attribute ButtonIcon.VB_Description = "Returns/sets a icon to be displayed in the menu button."

   Set ButtonIcon = m_Menus.Item(m_CurrentMenu).Icon

End Property

Public Property Let ButtonIcon(ByRef NewButtonIcon As StdPicture)

   Set ButtonIcon = NewButtonIcon

End Property

Public Property Set ButtonIcon(ByRef NewButtonIcon As StdPicture)

   Set m_Menus.Item(m_CurrentMenu).Icon = NewButtonIcon
   PropertyChanged "ButtonIcon"
   
   Call SetupCache

End Property

Public Property Get ButtonIconAlignment() As Alignments
Attribute ButtonIconAlignment.VB_Description = "Returns/sets the icon alignment of an menu button."

   ButtonIconAlignment = m_Menus.IconAlignment

End Property

Public Property Let ButtonIconAlignment(ByVal NewButtonIconAlignment As Alignments)

   m_Menus.IconAlignment = NewButtonIconAlignment
   PropertyChanged "ButtonIconAlignment"
   
   Call SetupCache

End Property

Public Property Get ButtonTag() As String
Attribute ButtonTag.VB_Description = "Stores any extra data needed for your program."

   ButtonTag = m_Menus.Item(m_CurrentMenu).Tag

End Property

Public Property Let ButtonTag(ByVal NewButtonTag As String)

   m_Menus.Item(m_CurrentMenu).Tag = NewButtonTag
   PropertyChanged "ButtonTag"

End Property

Public Property Get ButtonToolTipText() As String
Attribute ButtonToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the menu button."

   ButtonToolTipText = m_Menus.Item(m_CurrentMenu).ToolTipText

End Property

Public Property Let ButtonToolTipText(ByVal NewButtonToolTipText As String)

   m_Menus.Item(m_CurrentMenu).ToolTipText = NewButtonToolTipText
   PropertyChanged "ButtonToolTipText"

End Property

Public Property Get CurrentItem() As Integer
Attribute CurrentItem.VB_Description = "Returns/sets a value to select a menu item."

   CurrentItem = m_CurrentItem

End Property

Public Property Let CurrentItem(ByVal NewCurrentItem As Integer)

   If ShowMessage("The current Item", NewCurrentItem, m_MaxItems) Then Exit Property
   
   m_CurrentItem = NewCurrentItem

End Property

Public Property Get CurrentMenu() As Integer
Attribute CurrentMenu.VB_Description = "Returns/sets a value to select a menu."

   CurrentMenu = m_CurrentMenu

End Property

Public Property Let CurrentMenu(ByVal NewCurrentMenu As Integer)

   If ShowMessage("The current Menu", NewCurrentMenu, m_MaxMenus) Then Exit Property
   
   m_CurrentMenu = NewCurrentMenu
   m_CurrentItem = 1
   
   With m_Menus
      .CurrentMenu = m_CurrentMenu
      m_MaxItems = .Item(m_CurrentMenu).MenuItemCount
      ButtonCaption = .Item(m_CurrentMenu).Caption
   End With

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

   Set Font = picMenu.Font

End Property

Public Property Let Font(ByRef NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByRef NewFont As StdFont)

   With picMenu
      If NewFont Is Nothing Then Set NewFont = .Font
      If NewFont.Size >= HeightMenuButton - 10 Then NewFont.Size = HeightMenuButton - 10
      
      .Cls
      Set .Font = NewFont
      PropertyChanged "Font"
   End With
   
   Call SetupCache

End Property

Public Property Get FontBoldButtonCaption() As Boolean
Attribute FontBoldButtonCaption.VB_Description = "Determines whether the text for a menu button is bold."

   FontBoldButtonCaption = m_Menus.FontBoldButtonCaption

End Property

Public Property Let FontBoldButtonCaption(ByVal NewFontBoldButtonCaption As Boolean)

   m_Menus.FontBoldButtonCaption = NewFontBoldButtonCaption
   PropertyChanged "FontBoldButtonCaption"
   
   Call SetupCache

End Property

Public Property Get FontBoldItemCaption() As Boolean
Attribute FontBoldItemCaption.VB_Description = "Determines whether the text for a menu item is bold."

   FontBoldItemCaption = m_Menus.FontBoldItemCaption

End Property

Public Property Let FontBoldItemCaption(ByVal NewFontBoldItemCaption As Boolean)

   m_Menus.FontBoldItemCaption = NewFontBoldItemCaption
   PropertyChanged "FontBoldItemCaption"
   picMenu.Cls
   
   Call SetupCache

End Property

Public Property Get ItemAlignment() As AlignmentConstants
Attribute ItemAlignment.VB_Description = "Returns/sets the alignment of an menu item."

   ItemAlignment = m_ItemAlignment

End Property

Public Property Let ItemAlignment(ByRef NewItemAlignment As AlignmentConstants)

   m_ItemAlignment = NewItemAlignment
   m_Menus.ItemAlignment = m_ItemAlignment
   PropertyChanged "ItemAlignment"
   picMenu.Cls
   
   Call SetMenuObjectValue(m_ItemAlignment, MenuItemAlignment)
   Call SetupCache

End Property

Public Property Get ItemBackColor() As OLE_COLOR
Attribute ItemBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an menu item."

   ItemBackColor = picMenu.BackColor

End Property

Public Property Let ItemBackColor(ByVal NewItemBackColor As OLE_COLOR)

   picMenu.BackColor = NewItemBackColor
   picCache.BackColor = NewItemBackColor
   PropertyChanged "ItemBackColor"
   
   Call SetupCache

End Property

Public Property Get ItemCaption() As String
Attribute ItemCaption.VB_Description = "Returns/sets the text displayed in an menu item."

   ItemCaption = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Caption

End Property

Public Property Let ItemCaption(ByVal NewItemCaption As String)

   m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Caption = NewItemCaption
   PropertyChanged "ItemCaption"
   
   If Not Initializing Then
      picMenu.Cls
      
      Call picMenu_Paint
   End If

End Property

Public Property Get ItemForeColor() As OLE_COLOR
Attribute ItemForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an menu item."

   ItemForeColor = m_ItemForeColor

End Property

Public Property Let ItemForeColor(ByVal NewItemForeColor As OLE_COLOR)

   m_ItemForeColor = NewItemForeColor
   PropertyChanged "ItemForeColor"
   picMenu.Cls
   
   Call SetMenuObjectValue(m_ItemForeColor, MenuItemForeColor)
   Call picMenu_Paint

End Property

Public Property Get ItemIcon() As StdPicture
Attribute ItemIcon.VB_Description = "Returns/sets a icon to be displayed in the menu item."

   Set ItemIcon = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Icon

End Property

Public Property Let ItemIcon(ByRef NewItemIcon As StdPicture)

   Set ItemIcon = NewItemIcon

End Property

Public Property Set ItemIcon(ByRef NewItemIcon As StdPicture)

   If NewItemIcon Is Nothing Then Set NewItemIcon = m_ItemIcon
   
   Set m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Icon = NewItemIcon
   PropertyChanged "ItemIcon"
   
   Call SetupCache

End Property

Public Property Get ItemIconSize() As ItemIconSizeType
Attribute ItemIconSize.VB_Description = "Returns/sets the size of an icon object."

   ItemIconSize = m_ItemIconSize

End Property

Public Property Let ItemIconSize(ByVal NewIconSize As ItemIconSizeType)

   m_ItemIconSize = NewIconSize
   PropertyChanged "ItemIconSize"
   UserControl.Cls
   picMenu.Cls
   
   Call FitIcon

End Property

Public Property Get ItemKey() As String
Attribute ItemKey.VB_Description = "Returns/sets the value identifying a control in a control array."

   ItemKey = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Key

End Property

Public Property Let ItemKey(ByVal NewItemKey As String)

   m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Key = NewItemKey
   PropertyChanged "ItemKey"

End Property

Public Property Get ItemTag() As String
Attribute ItemTag.VB_Description = "Stores any extra data needed for your program."

   ItemTag = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Tag

End Property

Public Property Let ItemTag(ByVal NewItemTag As String)

   m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).Tag = NewItemTag
   PropertyChanged "ItemTag"

End Property

Public Property Get ItemToolTipText() As String
Attribute ItemToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the menu item."

   ItemToolTipText = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).ToolTipText

End Property

Public Property Let ItemToolTipText(ByVal NewItemToolTipText As String)

   m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).ToolTipText = NewItemToolTipText
   PropertyChanged "ItemToolTipText"

End Property

Public Property Get ItemType() As ItemTypes
Attribute ItemType.VB_Description = "Returns/sets a type used for the menu item."

   ItemType = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).ItemType

End Property

Public Property Let ItemType(ByVal NewItemType As ItemTypes)

   If (NewItemType <> CheckButton) And (NewItemType <> OptionButton) Then ItemValue = False
   
   m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).ItemType = NewItemType
   PropertyChanged "ItemType"

End Property

Public Property Get ItemValue() As Boolean

   ItemValue = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).ItemValue

End Property

Public Property Let ItemValue(ByVal NewItemValue As Boolean)

Dim itmType As ItemTypes

   itmType = m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).ItemType
   
   If (itmType <> CheckButton) And (itmType <> OptionButton) Then Exit Property
   
   m_Menus.Item(m_CurrentMenu).MenuItem(m_CurrentItem).ItemValue = NewItemValue
   PropertyChanged "ItemValue"
   
   Call SetupCache

End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets to lock/unlock an object to fasten the property changes."

   Locked = m_Locked

End Property

Public Property Let Locked(ByVal NewLocked As Boolean)

   m_Locked = NewLocked
   PropertyChanged "Locked"
   
   Call SetupCache

End Property

Public Property Get MaxItems() As Integer
Attribute MaxItems.VB_Description = "Returns/sets a maximum for the menu items."

   MaxItems = m_MaxItems

End Property

Public Property Let MaxItems(ByVal NewMaxItems As Integer)

Dim intCurrentItem As Integer

   If ShowMessage("MaxItems", NewMaxItems, MENU_MAX_ITEMS) Then Exit Property
   
   intCurrentItem = m_CurrentItem
   
   If NewMaxItems > m_MaxItems Then
      With m_Menus.Item(m_CurrentMenu)
         For m_CurrentItem = m_MaxItems + 1 To NewMaxItems
            .AddMenuItem MENU_ITEM_CAPTION & CStr(m_CurrentItem), m_CurrentItem, m_ItemIcon
            ItemCaption = MENU_ITEM_CAPTION & CStr(m_CurrentItem)
         Next 'm_CurrentItem
         
         m_CurrentItem = intCurrentItem
      End With
      
   ElseIf NewMaxItems < m_MaxItems Then
      With m_Menus.Item(m_CurrentMenu)
         For m_CurrentItem = m_MaxItems To NewMaxItems + 1 Step -1
            .DeleteMenuItem m_CurrentItem
         Next 'm_CurrentItem
         
         m_CurrentItem = intCurrentItem
         
         If NewMaxItems < m_CurrentItem Then m_CurrentItem = NewMaxItems
      End With
   End If
   
   m_MaxItems = NewMaxItems
   ItemAlignment = m_ItemAlignment
   ItemForeColor = m_ItemForeColor
   PropertyChanged "MaxItems"
   
   Call SetupCache

End Property

Public Property Get MaxMenus() As Integer
Attribute MaxMenus.VB_Description = "Returns/sets a maximum for the menus."

   MaxMenus = m_MaxMenus

End Property

Public Property Let MaxMenus(ByVal NewMaxMenus As Integer)

Dim intIndex    As Integer
Dim intCurrMenu As Integer

   If ShowMessage("MaxMenus", NewMaxMenus, MENU_MAX_MENUS) Then Exit Property
   If NewMaxMenus > 1 Then ButtonHideInSingleMenu = False
   
   If NewMaxMenus < m_StartupMenu Then
      ShowMessage "The startup Menu", m_StartupMenu, NewMaxMenus
      StartupMenu = 1
   End If
   
   Call UserControl_Resize
   
   If NewMaxMenus > m_MaxMenus Then
      intCurrMenu = m_CurrentMenu
      
      For m_CurrentMenu = m_MaxMenus + 1 To NewMaxMenus
         With m_Menus
            .Add "", m_CurrentMenu, picMenu
            ButtonCaption = MENU_BUTTON_CAPTION & CStr(m_CurrentMenu)
            Set .Item(m_CurrentMenu).ImageCache = picCache
            .Item(m_CurrentMenu).AddMenuItem MENU_ITEM_CAPTION & "1", 1, m_ItemIcon
         End With
      Next 'm_CurrentMenu
      
      m_CurrentMenu = intCurrMenu
      
   ElseIf NewMaxMenus < m_MaxMenus Then
      For intIndex = m_MaxMenus To NewMaxMenus + 1 Step -1
         With m_Menus
            Call .Delete(intIndex)
            
            If NewMaxMenus < m_CurrentMenu Then CurrentMenu = NewMaxMenus
         End With
      Next 'intIndex
   End If
   
   m_MaxMenus = NewMaxMenus
   m_Menus.NumberOfMenusChanged = True
   Appearance = m_Appearance
   ButtonBackColor = m_ButtonBackColor
   ButtonForeColor = m_ButtonForeColor
   ButtonGradientColor = m_ButtonGradientColor
   ButtonGradientType = m_ButtonGradientType
   ItemForeColor = m_ItemForeColor
   PropertyChanged "MaxMenus"
   
   Call SetupCache

End Property

Public Property Get MenuPassword() As Boolean
Attribute MenuPassword.VB_Description = "Determines whether a menu is protected by password."

   MenuPassword = m_Menus.Item(m_CurrentMenu).Password

End Property

Public Property Let MenuPassword(ByVal NewMenuPassword As Boolean)

   If NewMenuPassword And (m_CurrentMenu = m_StartupMenu) Then
      ShowMessage "PWD", m_StartupMenu
      NewMenuPassword = False
   End If
   
   m_Menus.Item(m_CurrentMenu).Password = NewMenuPassword
   PropertyChanged "MenuPassword"

End Property

Public Property Get OnlyFullItemsHit() As Boolean
Attribute OnlyFullItemsHit.VB_Description = "Determines whether an menu item can be hit only when it fits completely in the menu."

   OnlyFullItemsHit = m_Menus.OnlyFullItemsHit

End Property

Public Property Let OnlyFullItemsHit(ByVal NewOnlyFullItemsHit As Boolean)

   m_Menus.OnlyFullItemsHit = NewOnlyFullItemsHit
   PropertyChanged "OnlyFullItemsHit"
   
   Call SetupCache

End Property

Public Property Get OnlyFullItemsShow() As Boolean
Attribute OnlyFullItemsShow.VB_Description = "Determines whether an menu item can be showed only when it fits completely in the menu."

   OnlyFullItemsShow = m_Menus.OnlyFullItemsShow

End Property

Public Property Let OnlyFullItemsShow(ByVal NewOnlyFullItemsShow As Boolean)

   m_Menus.OnlyFullItemsShow = NewOnlyFullItemsShow
   PropertyChanged "OnlyFullItemsShow"
   
   Call SetupCache

End Property

Public Property Get SoundItemScroll() As Boolean
Attribute SoundItemScroll.VB_Description = "Determines whether a sound is played when the items will scrolled."

   SoundItemScroll = m_Menus.SoundItemScroll

End Property

Public Property Let SoundItemScroll(ByVal NewSoundItemScroll As Boolean)

   m_Menus.SoundItemScroll = NewSoundItemScroll
   PropertyChanged "SoundItemScroll"

End Property

Public Property Get SoundMenuOpen() As Boolean
Attribute SoundMenuOpen.VB_Description = "Determines whether a sound is played when a menu will open."

   SoundMenuOpen = m_SoundMenuOpen

End Property

Public Property Let SoundMenuOpen(ByVal NewSoundMenuOpen As Boolean)

   m_SoundMenuOpen = NewSoundMenuOpen
   PropertyChanged "SoundMenuOpen"

End Property

Public Property Get StartupMenu() As Integer
Attribute StartupMenu.VB_Description = "Returns/sets the menu for starting up the control."

   StartupMenu = m_StartupMenu

End Property

Public Property Let StartupMenu(ByVal NewStartupMenu As Integer)

   If ShowMessage("The startup Menu", NewStartupMenu, m_MaxMenus) Then Exit Property
   
   If m_Menus.Item(NewStartupMenu).Password Then
      ShowMessage "START", NewStartupMenu
      Exit Property
   End If
   
   m_StartupMenu = NewStartupMenu
   PropertyChanged "StartupMenu"

End Property

Public Function GetMenuItems(ByVal SelectedMenu As Integer) As Integer

   If SelectedMenu < 1 Or SelectedMenu > m_MaxMenus Then Exit Function
   
   GetMenuItems = m_Menus.MenuItems(SelectedMenu)

End Function

Public Sub MoveToCurrentItem()

Dim intItemsShowed As Integer
Dim intMoves       As Integer
Dim intMenuItems   As Integer
Dim intTopItem     As Integer

   If Not picMenu.Visible Then
      ToCurrentItem = True
      Exit Sub
   End If
   
   ToCurrentItem = False
   intItemsShowed = m_Menus.ItemsShowed
   intTopItem = m_Menus.TopItem
   intMenuItems = m_Menus.MenuItems(m_CurrentMenu)
   
   If m_CurrentItem < intTopItem Then
      intMoves = m_CurrentItem - intTopItem
      
   ElseIf m_CurrentItem > intTopItem + intItemsShowed - 1 Then
      intMoves = intTopItem + (m_CurrentItem - intItemsShowed - 1)
   End If
   
   If intMoves Then Call m_Menus.MoveToItem(intMoves)

End Sub

Private Function GetMenuButtonHeight(ByVal Height As ButtonHeights) As Long

Const MENU_WINDOW_SPACE As Long = 33

   If Height = -1 Then
      GetMenuButtonHeight = (HeightMenuButton * 2) + MENU_WINDOW_SPACE
      
   Else
      GetMenuButtonHeight = MENU_BUTTON_MIN_HEIGHT + (16 And Height = High)
   End If

End Function

Private Function ShowMessage(ByVal Message As String, ByVal SelectedMenu As Integer, Optional ByVal MaxMenus As Integer) As Long

   If (SelectedMenu < 1) Or (SelectedMenu > MaxMenus) Then
      Beep
      
      If Message = "PWD" Then
         Message = "Startup Menu (" & SelectedMenu & ") can't be protect by password!"
         
      ElseIf Message = "START" Then
         Message = "Menu (" & SelectedMenu & ") can't be the startup menu, because it's protected by password!"
         
      Else
         Message = Message & " is between 1 and " & MaxMenus
      End If
      
      ShowMessage = MsgBox(Message, vbOKOnly)
   End If

End Function

Public Sub Refresh()

   Call FitIcon

End Sub

Private Sub DrawCacheMenuButton()

Dim intPosition As Integer
Dim rctMenu     As Rect

   If m_Menus.ButtonHideInSingleMenu Then Exit Sub
   
   On Local Error Resume Next
   
   With rctMenu
      .Left = 0
      .Top = 0
      .Right = picCache.ScaleWidth
      .Bottom = HeightMenuButton
   End With
   
   Call DrawGradient(rctMenu, picCache, m_ButtonGradientType, m_ButtonGradientColor, m_ButtonBackColor)
   
   intPosition = m_Appearance + 1
   DrawEdge picCache.hDC, rctMenu, m_MenuBorder, BF_RECT
   
   If m_ButtonGradientType = NoGradient Then picCache.Line (intPosition, intPosition)-(picCache.ScaleWidth - 2, HeightMenuButton - (1 + intPosition)), m_ButtonBackColor, BF
   
   On Local Error GoTo 0

End Sub

Private Sub FitIcon()

   SizeIcon = (m_ItemIconSize + 1) * 16
   m_Menus.ItemIconSize = SizeIcon
   
   Call SetupCache

End Sub

Private Sub ProcessDefaultIcon()

   If m_ItemIcon Is Nothing Then Set m_ItemIcon = UserControl.Picture
   
   UserControl.Picture = LoadPicture()

End Sub

Private Sub ResetHitObject()

   MenuHitObject = 0
   MenuHitType = 0
   MenuObjectX = 0
   MenuObjectY = 0

End Sub

Private Sub SetMenuObjectValue(ByVal NewValue As Long, ByVal MenuObject As MenuObjects)

Dim intCount As Integer
Dim intItem  As Integer

   For intCount = 1 To m_Menus.Count
      With m_Menus.Item(intCount)
         If MenuObject = MenuButtonBackColor Then
            .BackColor = NewValue
            
         ElseIf MenuObject = MenuBorder Then
            .Border = NewValue
            
         ElseIf MenuObject = MenuButtonHeight Then
            .ButtonHeight = NewValue
            
         ElseIf MenuObject = MenuButtonForeColor Then
            .ForeColor = NewValue
            
         ElseIf MenuObject = MenuButtonGradientColor Then
            .GradientColor = NewValue
            
         ElseIf MenuObject = MenuButtonGradientType Then
            .GradientType = NewValue
            
         ElseIf MenuObject = MenuItemAlignment Then
            For intItem = 1 To .MenuItemCount
               .MenuItem(intItem).ItemAlignment = NewValue
            Next 'intItem
            
         ElseIf MenuObject = MenuItemForeColor Then
            For intItem = 1 To .MenuItemCount
               .MenuItem(intItem).ItemForeColor = NewValue
            Next 'intItem
         End If
      End With
   Next 'intCount

End Sub

Private Sub SetupCache()

Dim intIcon   As Integer
Dim intItem   As Integer
Dim intMenu   As Integer
Dim lngOffset As Long

   If Initializing Or m_Locked Then Exit Sub
   
   On Local Error Resume Next
   
   With picCache
      .Cls
      lngOffset = HeightMenuButton * 2 + SizeIcon
      .Height = HeightMenuButton * 2 + (m_Menus.TotalMenuItems + 1) * SizeIcon
      
      Call DrawCacheMenuButton
      
      For intMenu = 1 To m_Menus.Count
         For intItem = 1 To m_Menus.Item(intMenu).MenuItemCount
            DrawIconEx .hDC, 0, lngOffset + intIcon * SizeIcon, m_Menus.Item(intMenu).MenuItem(intItem).Icon.Handle, SizeIcon, SizeIcon, 0, 0, DI_NORMAL
            intIcon = intIcon + 1
         Next 'intItem
      Next 'intMenu
   End With
   
   On Local Error GoTo 0
   
   Call picMenu_Paint

End Sub

Private Sub picCache_Resize()

   Call DrawCacheMenuButton

End Sub

' for suporting picMenu double click
Private Sub picMenu_DblClick()

Dim ptaMouse As PointAPI

  GetCursorPos ptaMouse
  ScreenToClient picMenu.hWnd, ptaMouse
  
  Call picMenu_MouseDown(vbLeftButton, 0, CSng(ptaMouse.X), CSng(ptaMouse.Y))

End Sub

Private Sub picMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then
      With m_Menus
         MenuObjectX = CLng(X)
         MenuObjectY = CLng(Y)
         MenuHitObject = .MouseProcess(MOUSE_DOWN, MenuObjectX, MenuObjectY, MenuHitType)
      End With
   End If

End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim intHitType As Integer

   With m_Menus
      If Button = vbLeftButton Then
         hWndMenuBar = 0
         tmrMouseOut.Enabled = False
         
         If MenuHitType Then Exit Sub
         If .MouseProcess(MOUSE_CHECK, CLng(X), CLng(Y)) + MenuHitObject = 0 Then Exit Sub
         
      Else
         hWndMenuBar = picMenu.hWnd
         tmrMouseOut.Enabled = True
         .MouseProcess MOUSE_MOVE, CLng(X), CLng(Y)
         MenuHitObject = 0
         MenuHitType = 0
      End If
   End With

End Sub

Private Sub picMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Const SOUND_MENU_CLICKED As Integer = 1

Static blnArrow          As Boolean
Static intPrevOption     As Integer

Dim blnCancel            As Boolean
Dim blnCheckButton       As Boolean
Dim blnItemValue         As Boolean
Dim intCurrentMenu       As Integer
Dim intHitType           As Integer
Dim intIndex             As Integer

   If Button = vbLeftButton Then
      With m_Menus
         If MenuHitType = HIT_TYPE_ARROW Then
            intIndex = .MouseProcess(MOUSE_DOWN, CLng(X), CLng(Y), intHitType)
            
            If intHitType <> HIT_TYPE_ARROW Then
               If intHitType = HIT_TYPE_MENU_BUTTON Then picMenu.Refresh
               
               .MouseProcess MOUSE_MOVE, CLng(X), CLng(Y)
               Exit Sub
               
            ElseIf intIndex = HIT_TYPE_ARROW And ((CLng(Y) > MenuObjectY + ARROW_BUTTON_SIZE) Or (CLng(Y) < MenuObjectY - ARROW_BUTTON_SIZE)) Then
               m_Menus.Item(m_CurrentMenu).MouseProcessForArrows MOUSE_UP, CLng(X), CLng(Y)
               
               Call ResetHitObject
               
               Exit Sub
            End If
            
         Else
            intIndex = .MouseProcess(MOUSE_CHECK, CLng(X), CLng(Y), intHitType)
            
            If intIndex = HIT_TYPE_ARROW And intHitType = HIT_TYPE_ARROW Then blnArrow = True
         End If
         
         If blnArrow Or MenuHitObject Then
            Call .RestoreButton(MOUSE_UP, MenuObjectX, MenuObjectY)
            
            If blnArrow Or (intHitType = HIT_TYPE_MENU_ITEM) Then
               Call ResetHitObject
               
               blnArrow = False
               Exit Sub
            End If
         End If
         
         If (intIndex = MenuHitObject) And (intHitType = HIT_TYPE_MENU_BUTTON) Then
            RaiseEvent BeforeOpenMenu(intIndex, CurrentMenu = intIndex, blnCancel)
            
            If Not blnCancel And (CurrentMenu <> intIndex) Then
               If .Item(intIndex).Password Then
                  RaiseEvent CheckMenuWithPassword(intIndex, blnCancel)
                  
                  If Not blnCancel Then
                     CurrentMenu = intIndex
                     
                     If m_SoundMenuOpen Then Call PlaySound(SOUND_MENU_CLICKED)
                  End If
                  
               Else
                  CurrentMenu = intIndex
                  
                  If m_SoundMenuOpen Then Call PlaySound(SOUND_MENU_CLICKED)
               End If
               
               If intIndex = CurrentMenu Then RaiseEvent MenuClick(CurrentMenu)
            End If
            
            If MenuHitType = HIT_TYPE_MENU_ITEM Then .MouseProcess MOUSE_MOVE, MenuObjectX, MenuObjectY
            
            Call ResetHitObject
            
            Exit Sub
         End If
         
         intCurrentMenu = m_CurrentMenu
         intIndex = .MouseProcess(MOUSE_UP, CLng(X), CLng(Y), intHitType)
         
         If intIndex And (intHitType = HIT_TYPE_MENU_ITEM) Then
            With .Item(m_CurrentMenu).MenuItem(intIndex)
               If .ItemType = ResetButton Then
                  With m_Menus.Item(m_CurrentMenu)
                     For intIndex = 1 To .MenuItemCount
                        If (.MenuItem(intIndex).ItemType = CheckButton) And (.MenuItem(intIndex).ItemValue = True) Then
                           .MenuItem(intIndex).ItemValue = False
                           blnCheckButton = True
                        End If
                     Next 'intItem
                  End With
                  
                  If blnCheckButton Then
                     blnCheckButton = False
                     
                  Else
                     If intPrevOption Then m_Menus.Item(m_CurrentMenu).MenuItem(intPrevOption).ItemValue = False
                     
                     intPrevOption = 0
                  End If
                  
                  RaiseEvent MenuItemClick(m_CurrentMenu, .Index, .Key, .ItemType, blnItemValue, 0)
                  
               ElseIf .ItemType = CheckButton Then
                  m_Menus.Item(m_CurrentMenu).MenuItem(intIndex).ItemValue = Not m_Menus.Item(m_CurrentMenu).MenuItem(intIndex).ItemValue
                  
                  RaiseEvent MenuItemClick(m_CurrentMenu, .Index, .Key, .ItemType, .ItemValue, 0)
                  
               ElseIf .ItemType = OptionButton Then
                  blnItemValue = Not .ItemValue
                  
                  If intPrevOption Then m_Menus.Item(m_CurrentMenu).MenuItem(intPrevOption).ItemValue = False
                  
                  .ItemValue = blnItemValue
                  RaiseEvent MenuItemClick(m_CurrentMenu, .Index, .Key, .ItemType, .ItemValue, intPrevOption)
                  intPrevOption = .Index
                  
               ElseIf (.ItemType = LockMenuButton) And m_Menus.Item(m_CurrentMenu).Password Then
                  RaiseEvent MenuItemClick(m_CurrentMenu, .Index, .Key, .ItemType, blnItemValue, 0)
                  RaiseEvent LockMenuWithPassword(m_CurrentMenu, blnCancel)
                  
                  If Not blnCancel Then
                     m_Menus.MouseProcess MOUSE_MOVE, 0, 0
                     CurrentMenu = m_StartupMenu
                     
                     If m_SoundMenuOpen Then Call PlaySound(SOUND_MENU_CLICKED)
                     
                     RaiseEvent MenuClick(m_CurrentMenu)
                  End If
                  
               ' DefaultButton
               Else
                  RaiseEvent MenuItemClick(m_CurrentMenu, .Index, .Key, .ItemType, 0, 0)
               End If
               
               If intCurrentMenu = m_CurrentMenu Then .HitTest MOUSE_MOVE, 0, 0
            End With
         End If
         
         .MouseProcess MOUSE_MOVE, CLng(X), CLng(Y)
         
         Call ResetHitObject
      End With
   End If

End Sub

Private Sub picMenu_Paint()

   If picMenu.Visible Then
      Call m_Menus.Paint
      
      If ToCurrentItem Then Call MoveToCurrentItem
   End If

End Sub

Private Sub tmrMouseOut_Timer()

Dim ptaMouse As PointAPI

   GetCursorPos ptaMouse
   
   If hWndMenuBar And (WindowFromPoint(ptaMouse.X, ptaMouse.Y) <> hWndMenuBar) Then
      Call picMenu_MouseMove(0, 0, 0, 0)
      
      hWndMenuBar = 0
   End If

End Sub

Private Sub UserControl_Initialize()

   Set m_Menus = New clsMenus
   Set m_Menus.Menu = picMenu
   Set m_Menus.Cache = picCache

End Sub

Private Sub UserControl_InitProperties()

   Initializing = True
   m_Menus.Animation = True
   m_Appearance = [3D]
   m_BorderStyle = [Fixed Single]
   HeightMenuButton = GetMenuButtonHeight(Low)
   m_Menus.ButtonHeight = HeightMenuButton
   m_ButtonBackColor = vbButtonFace
   m_Menus.CaptionAlignment = vbCenter
   m_ButtonForeColor = vbButtonText
   m_ButtonGradientColor = vbButtonText
   m_ItemForeColor = vbHighlightText
   m_MenuBorder = BDR_RAISED
   m_ItemIconSize = [32x32]
   
   Call ProcessDefaultIcon
   
   With picCache
      .Width = picMenu.Width
      .Height = GetMenuButtonHeight(-1)
      .BackColor = vbApplicationWorkspace
      picMenu.BackColor = .BackColor
   End With
   
   MaxMenus = 1
   CurrentMenu = 1
   StartupMenu = 1
   SoundMenuOpen = True
   SoundItemScroll = True
   ButtonGradientColor = vbButtonFace
   ButtonGradientType = NoGradient
   Font = Ambient.Font
   ItemAlignment = vbCenter
   Initializing = False
   
   Call FitIcon

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim intItem As Integer
Dim intMenu As Integer
Dim strItem As String
Dim strMenu As String

   With PropBag
      Initializing = True
      Animation = .ReadProperty("Animation", True)
      m_Appearance = .ReadProperty("Appearance", [3D])
      
      If m_Appearance = [3D] Then
         m_MenuBorder = BDR_RAISED
         
      Else
         m_MenuBorder = BDR_RAISEDINNER
      End If
      
      m_BorderStyle = .ReadProperty("BorderStyle", [Fixed Single])
      
      If m_BorderStyle < Edged Then
         UserControl.BorderStyle = m_BorderStyle
         
      Else
         UserControl.BorderStyle = None
      End If
      
      m_ButtonBackColor = .ReadProperty("ButtonBackColor", vbButtonFace)
      m_Menus.CaptionAlignment = .ReadProperty("ButtonCaptionAlignment", vbCenter)
      m_ButtonForeColor = .ReadProperty("ButtonForeColor", vbButtonText)
      m_ButtonGradientColor = .ReadProperty("ButtonGradientColor", m_ButtonBackColor)
      m_ButtonGradientType = .ReadProperty("ButtonGradientType", NoGradient)
      m_ButtonHeight = .ReadProperty("ButtonHeight", Low)
      m_Menus.ButtonHideInSingleMenu = .ReadProperty("ButtonHideInSingleMenu", False)
      HeightMenuButton = GetMenuButtonHeight(m_ButtonHeight)
      m_Menus.ButtonHeight = HeightMenuButton
      m_Menus.IconAlignment = .ReadProperty("ButtonIconAlignment", [Left Justify])
      Font = .ReadProperty("Font")
      m_Menus.FontBoldButtonCaption = .ReadProperty("FontBoldButtonCaption", True)
      m_Menus.FontBoldItemCaption = .ReadProperty("FontBoldItemCaption", True)
      ItemAlignment = .ReadProperty("ItemAlignment", vbCenter)
      picMenu.BackColor = .ReadProperty("ItemBackColor", vbApplicationWorkspace)
      ItemForeColor = .ReadProperty("ItemForeColor", vbHighlightText)
      Set m_ItemIcon = .ReadProperty("ItemIcon", Nothing)
      m_ItemIconSize = .ReadProperty("ItemIconSize", [32x32])
      m_Locked = .ReadProperty("Locked", False)
      MaxMenus = .ReadProperty("MaxMenus", 1)
      
      With picCache
         .BackColor = picMenu.BackColor
         .Width = UserControl.Width
         .Height = GetMenuButtonHeight(-1)
      End With
      
      Call ProcessDefaultIcon
      
      For intMenu = 1 To m_MaxMenus
         CurrentMenu = intMenu
         strMenu = CStr(intMenu)
         ButtonCaption = .ReadProperty("ButtonCaption" & strMenu, MENU_BUTTON_CAPTION & strMenu)
         ButtonIcon = .ReadProperty("ButtonIcon" & strMenu, Nothing)
         ButtonTag = .ReadProperty("ButtonTag" & strMenu, "")
         ButtonToolTipText = .ReadProperty("ButtonToolTipText" & strMenu, "")
         MaxItems = .ReadProperty("MaxItems" & strMenu, 1)
         MenuPassword = .ReadProperty("MenuPassword" & strMenu, False)
         
         For intItem = 1 To m_Menus.Item(intMenu).MenuItemCount
            strItem = strMenu & "_" & CStr(intItem)
            CurrentItem = intItem
            ItemCaption = .ReadProperty("ItemCaption" & strItem, MENU_ITEM_CAPTION & CStr(intItem))
            Set ItemIcon = .ReadProperty("ItemIcon" & strItem, m_ItemIcon)
            ItemKey = .ReadProperty("ItemKey" & strItem, "")
            ItemTag = .ReadProperty("ItemTag" & strItem, "")
            ItemToolTipText = .ReadProperty("ItemToolTipText" & strItem, "")
            ItemType = .ReadProperty("ItemType" & strItem, DefaultButton)
            ItemValue = .ReadProperty("ItemValue" & strItem, False)
         Next 'intItem
      Next 'intMenu
      
      CurrentItem = 1
      OnlyFullItemsHit = .ReadProperty("OnlyFullItemsHit", False)
      OnlyFullItemsShow = .ReadProperty("OnlyFullItemsShow", False)
      m_SoundMenuOpen = .ReadProperty("SoundMenuOpen", True)
      m_Menus.SoundItemScroll = .ReadProperty("SoundItemScroll", True)
      StartupMenu = .ReadProperty("StartupMenu", 1)
      CurrentMenu = m_StartupMenu
   End With
   
   For intMenu = 1 To m_Menus.Count
      With m_Menus.Item(intMenu)
         .BackColor = m_ButtonBackColor
         .Border = m_MenuBorder
         .ForeColor = m_ButtonForeColor
         .GradientColor = m_ButtonGradientColor
         .GradientType = m_ButtonGradientType
      End With
   Next 'intMenu
   
   If Ambient.UserMode Then tmrMouseOut.Enabled = True
   
   Initializing = False
   
   Call FitIcon

End Sub

Private Sub UserControl_Resize()

Static blnBusy As Boolean

Dim lngHeight  As Long
Dim rctControl As Rect

   If blnBusy Then Exit Sub
   
   blnBusy = True
   lngHeight = 852 + ((m_MaxMenus - m_Menus.ButtonHideInSingleMenu) * HeightMenuButton) * Screen.TwipsPerPixelY
   
   If Height < lngHeight Then Height = lngHeight
   If Width < 732 Then Width = 732
   
   With picMenu
      .Top = 0 + (1 And (m_BorderStyle = Edged))
      .Left = 0 + (1 And (m_BorderStyle = Edged))
      .Width = UserControl.ScaleWidth - (2 And (m_BorderStyle = Edged))
      .Height = UserControl.ScaleHeight - (2 And (m_BorderStyle = Edged))
      picCache.Width = .Width
      picCache.Height = GetMenuButtonHeight(-1)
      .Refresh
   End With
   
   If m_BorderStyle = Edged Then
      With rctControl
         .Top = 0
         .Left = 0
         .Right = UserControl.ScaleWidth
         .Bottom = UserControl.ScaleHeight
         DrawEdge hDC, rctControl, BDR_SUNKENOUTER, BF_RECT
      End With
   End If
   
   blnBusy = False

End Sub

Private Sub UserControl_Terminate()

   tmrMouseOut.Enabled = False
   Set m_MenuButtonIcon = Nothing
   Set m_ItemIcon = Nothing
   Set m_Menus = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Dim intCurrItem As Integer
Dim intCurrMenu As Integer
Dim intItem     As Integer
Dim intMenu     As Integer
Dim strItem     As String
Dim strMenu     As String

   With PropBag
      Initializing = True
      .WriteProperty "Animation", Animation, True
      .WriteProperty "Appearance", m_Appearance, [3D]
      .WriteProperty "BorderStyle", m_BorderStyle, [Fixed Single]
      .WriteProperty "ButtonBackColor", m_ButtonBackColor, vbButtonFace
      .WriteProperty "ButtonCaptionAlignment", m_Menus.CaptionAlignment, vbCenter
      .WriteProperty "ButtonForeColor", m_ButtonForeColor, vbButtonText
      .WriteProperty "ButtonGradientColor", m_ButtonGradientColor, m_ButtonBackColor
      .WriteProperty "ButtonGradientType", m_ButtonGradientType, NoGradient
      .WriteProperty "ButtonHeight", m_ButtonHeight, Low
      .WriteProperty "ButtonHideInSingleMenu", m_Menus.ButtonHideInSingleMenu, False
      .WriteProperty "ButtonIconAlignment", m_Menus.IconAlignment, [Left Justify]
      .WriteProperty "Font", picMenu.Font, Ambient.Font
      .WriteProperty "FontBoldButtonCaption", m_Menus.FontBoldButtonCaption, True
      .WriteProperty "FontBoldItemCaption", m_Menus.FontBoldItemCaption, True
      .WriteProperty "ItemAlignment", m_ItemAlignment, vbCenter
      .WriteProperty "ItemBackColor", picMenu.BackColor, vbApplicationWorkspace
      .WriteProperty "ItemForeColor", m_ItemForeColor, vbHighlightText
      .WriteProperty "ItemIcon", m_ItemIcon, m_ItemIcon
      .WriteProperty "ItemIconSize", m_ItemIconSize, [32x32]
      .WriteProperty "Locked", m_Locked, False
      .WriteProperty "MaxMenus", m_MaxMenus, 1
      intCurrMenu = CurrentMenu
      intCurrItem = CurrentItem
      
      For intMenu = 1 To m_MaxMenus
         strMenu = CStr(intMenu)
         .WriteProperty "ButtonCaption" & strMenu, m_Menus.Item(intMenu).Caption, MENU_BUTTON_CAPTION & strMenu
         .WriteProperty "ButtonIcon" & strMenu, m_Menus.Item(intMenu).Icon, Nothing
         .WriteProperty "ButtonTag" & strMenu, m_Menus.Item(intMenu).Tag, ""
         .WriteProperty "ButtonToolTipText" & strMenu, m_Menus.Item(intMenu).ToolTipText, ""
         .WriteProperty "MaxItems" & strMenu, m_Menus.Item(intMenu).MenuItemCount, 1
         .WriteProperty "MenuPassword" & strMenu, m_Menus.Item(intMenu).Password, False
         CurrentMenu = intMenu
         
         For intItem = 1 To m_Menus.Item(intMenu).MenuItemCount
            CurrentItem = intItem
            strItem = strMenu & "_" & CStr(intItem)
            .WriteProperty "ItemCaption" & strItem, ItemCaption, MENU_ITEM_CAPTION & CStr(intItem)
            .WriteProperty "ItemIcon" & strItem, ItemIcon, Nothing
            .WriteProperty "ItemKey" & strItem, ItemKey, ""
            .WriteProperty "ItemTag" & strItem, ItemTag, ""
            .WriteProperty "ItemToolTipText" & strItem, ItemToolTipText, ""
            .WriteProperty "ItemType" & strItem, ItemType, DefaultButton
            .WriteProperty "ItemValue" & strItem, ItemValue, False
         Next 'intItem
      Next 'intMenu
      
      CurrentMenu = intCurrMenu
      CurrentItem = intCurrItem
      .WriteProperty "OnlyFullItemsHit", m_Menus.OnlyFullItemsHit, False
      .WriteProperty "OnlyFullItemsShow", m_Menus.OnlyFullItemsShow, False
      .WriteProperty "SoundMenuOpen", m_SoundMenuOpen, True
      .WriteProperty "SoundItemScroll", m_Menus.SoundItemScroll, True
      .WriteProperty "StartupMenu", m_StartupMenu, 1
      Initializing = False
   End With

End Sub

