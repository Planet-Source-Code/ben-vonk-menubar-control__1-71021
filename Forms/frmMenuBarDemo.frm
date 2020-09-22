VERSION 5.00
Object = "*\A..\MenuBarOcx.vbp"
Begin VB.Form frmMenuBarDemo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MenuBar (Demo)"
   ClientHeight    =   5412
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7344
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenuBarDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5412
   ScaleWidth      =   7344
   StartUpPosition =   2  'CenterScreen
   Begin MenuBarOcx.MenuBar mnbShow 
      Height          =   3132
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   5525
      ButtonBackColor =   6291591
      ButtonCaptionAlignment=   1
      ButtonForeColor =   14737632
      ButtonGradientColor=   14221311
      ButtonGradientType=   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBoldItemCaption=   0   'False
      ItemAlignment   =   1
      ItemBackColor   =   14221311
      ItemForeColor   =   12582912
      ItemIconSize    =   0
      MaxMenus        =   3
      MaxItems1       =   4
      ItemIcon1_1     =   "frmMenuBarDemo.frx":014A
      ItemIcon1_2     =   "frmMenuBarDemo.frx":0464
      ItemIcon1_3     =   "frmMenuBarDemo.frx":077E
      ItemIcon1_4     =   "frmMenuBarDemo.frx":0A98
      MaxItems2       =   5
      ItemIcon2_1     =   "frmMenuBarDemo.frx":0DB2
      ItemIcon2_2     =   "frmMenuBarDemo.frx":10CC
      ItemIcon2_3     =   "frmMenuBarDemo.frx":13E6
      ItemIcon2_4     =   "frmMenuBarDemo.frx":1700
      ItemIcon2_5     =   "frmMenuBarDemo.frx":1A1A
      MaxItems3       =   3
      ItemIcon3_1     =   "frmMenuBarDemo.frx":1D34
      ItemIcon3_2     =   "frmMenuBarDemo.frx":204E
      ItemIcon3_3     =   "frmMenuBarDemo.frx":2368
   End
   Begin MenuBarOcx.MenuBar mnbDemo 
      Height          =   3132
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   5525
      BorderStyle     =   2
      ButtonBackColor =   12582912
      ButtonForeColor =   65535
      ButtonGradientColor=   15132390
      ButtonGradientType=   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemAlignment   =   0
      ItemBackColor   =   16773862
      ItemForeColor   =   12582912
      MaxMenus        =   4
      ButtonCaption1  =   "Default"
      ButtonToolTipText1=   "Menu with default buttons"
      MaxItems1       =   3
      ItemIcon1_1     =   "frmMenuBarDemo.frx":2682
      ItemToolTipText1_1=   "DefaultButton"
      ItemIcon1_2     =   "frmMenuBarDemo.frx":299C
      ItemToolTipText1_2=   "DefaultButton"
      ItemIcon1_3     =   "frmMenuBarDemo.frx":2CB6
      ItemToolTipText1_3=   "DefaultButton"
      ButtonCaption2  =   "Check"
      ButtonToolTipText2=   "Menu with check buttons"
      MaxItems2       =   4
      ItemIcon2_1     =   "frmMenuBarDemo.frx":2FD0
      ItemToolTipText2_1=   "CheckButton"
      ItemType2_1     =   1
      ItemIcon2_2     =   "frmMenuBarDemo.frx":32EA
      ItemToolTipText2_2=   "CheckButton"
      ItemType2_2     =   1
      ItemIcon2_3     =   "frmMenuBarDemo.frx":3604
      ItemToolTipText2_3=   "CheckButton"
      ItemType2_3     =   1
      ItemIcon2_4     =   "frmMenuBarDemo.frx":391E
      ItemToolTipText2_4=   "ResetButton"
      ItemType2_4     =   3
      ButtonCaption3  =   "Option"
      ButtonToolTipText3=   "Menu with option buttons"
      MaxItems3       =   4
      ItemIcon3_1     =   "frmMenuBarDemo.frx":3C38
      ItemToolTipText3_1=   "OptionButton"
      ItemType3_1     =   2
      ItemIcon3_2     =   "frmMenuBarDemo.frx":3F52
      ItemToolTipText3_2=   "OptionButton"
      ItemType3_2     =   2
      ItemIcon3_3     =   "frmMenuBarDemo.frx":426C
      ItemToolTipText3_3=   "OptionButton"
      ItemType3_3     =   2
      ItemIcon3_4     =   "frmMenuBarDemo.frx":4586
      ItemToolTipText3_4=   "ResetButton"
      ItemType3_4     =   3
      ButtonCaption4  =   "Password"
      ButtonToolTipText4=   "Menu password protect"
      MaxItems4       =   2
      MenuPassword4   =   -1  'True
      ItemIcon4_1     =   "frmMenuBarDemo.frx":48A0
      ItemToolTipText4_1=   "DefaultButton"
      ItemIcon4_2     =   "frmMenuBarDemo.frx":4BBA
      ItemToolTipText4_2=   "LockMenuButton"
      ItemType4_2     =   4
   End
   Begin MenuBarOcx.MenuBar mnbMenu 
      Height          =   5160
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   9102
      BorderStyle     =   2
      ButtonBackColor =   12632256
      ButtonCaptionAlignment=   0
      ButtonForeColor =   9527808
      ButtonGradientColor=   15000804
      ButtonGradientType=   4
      ButtonHeight    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemBackColor   =   -2147483633
      ItemForeColor   =   12582912
      ItemIcon1_1     =   "frmMenuBarDemo.frx":4ED4
      OnlyFullItemsHit=   -1  'True
   End
   Begin VB.Frame fraResult 
      Caption         =   "Result"
      ForeColor       =   &H00800080&
      Height          =   1812
      Left            =   3000
      TabIndex        =   3
      Top             =   3480
      Width           =   4212
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3972
      End
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   384
      Index           =   1
      Left            =   4680
      TabIndex        =   6
      Top             =   3240
      Width           =   204
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   384
      Index           =   0
      Left            =   3240
      TabIndex        =   5
      Top             =   3240
      Width           =   204
   End
   Begin VB.Image imgMenu 
      Height          =   384
      Index           =   2
      Left            =   7680
      Picture         =   "frmMenuBarDemo.frx":51EE
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMenu 
      Height          =   384
      Index           =   1
      Left            =   7680
      Picture         =   "frmMenuBarDemo.frx":5AB8
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMenu 
      Height          =   384
      Index           =   0
      Left            =   7680
      Picture         =   "frmMenuBarDemo.frx":6782
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgTools 
      Height          =   384
      Index           =   5
      Left            =   9480
      Picture         =   "frmMenuBarDemo.frx":704C
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgTools 
      Height          =   384
      Index           =   4
      Left            =   9480
      Picture         =   "frmMenuBarDemo.frx":7916
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgTools 
      Height          =   384
      Index           =   3
      Left            =   9480
      Picture         =   "frmMenuBarDemo.frx":81E0
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgTools 
      Height          =   384
      Index           =   2
      Left            =   9480
      Picture         =   "frmMenuBarDemo.frx":8AAA
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgTools 
      Height          =   384
      Index           =   1
      Left            =   9480
      Picture         =   "frmMenuBarDemo.frx":9374
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgTools 
      Height          =   384
      Index           =   0
      Left            =   9480
      Picture         =   "frmMenuBarDemo.frx":9C3E
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   8
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":A508
      Top             =   3960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   7
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":ADD2
      Top             =   3480
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   6
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":B69C
      Top             =   3000
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   5
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":BF66
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   4
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":C830
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   3
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":D0FA
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   2
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":D9C4
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   1
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":E28E
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings 
      Height          =   384
      Index           =   0
      Left            =   8880
      Picture         =   "frmMenuBarDemo.frx":EB58
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgAdmin 
      Height          =   384
      Index           =   3
      Left            =   8280
      Picture         =   "frmMenuBarDemo.frx":F422
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgAdmin 
      Height          =   384
      Index           =   2
      Left            =   8280
      Picture         =   "frmMenuBarDemo.frx":FCEC
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgAdmin 
      Height          =   384
      Index           =   1
      Left            =   8280
      Picture         =   "frmMenuBarDemo.frx":105B6
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgAdmin 
      Height          =   384
      Index           =   0
      Left            =   8280
      Picture         =   "frmMenuBarDemo.frx":10E80
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmMenuBarDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EndDemo()

   End

End Sub

Private Sub FillMenus()

Dim intItem As Integer
Dim intMenu As Integer

   With mnbMenu
      .Locked = True
      .MaxMenus = 3
      .CurrentMenu = .MaxMenus
      
      For intMenu = 1 To .MaxMenus
         .CurrentMenu = intMenu
         .MaxItems = Choose(intMenu, 4, 9, 6)
         .ButtonCaption = Choose(intMenu, "Modules", "Settings", "Tools")
         .ButtonIcon = imgMenu.Item(intMenu - 1).Picture
         
         For intItem = 1 To .MaxItems
            .CurrentItem = intItem
            
            If intMenu = 1 Then
               .ItemCaption = Choose(intItem, "Add", "Message", "Picture", "Exit")
               .ItemIcon = imgAdmin.Item(intItem - 1).Picture
               
            ElseIf intMenu = 2 Then
               .ItemCaption = Choose(intItem, "Date", "Time", "Wipe info", "Favorits", "Agenda", "First quarter", "Second quarter", "Third quarter", "Fourth quarter")
               .ItemIcon = imgSettings.Item(intItem - 1).Picture
               
            ElseIf intMenu = 3 Then
               .ItemCaption = Choose(intItem, "Disk", "Trash", "Refresh", "Global info", "Settings", "Lock menu")
               .ItemIcon = imgTools.Item(intItem - 1).Picture
            End If
         Next 'intItem
      Next 'intMenu
      
      .MenuPassword = True
      .CurrentMenu = .StartupMenu
      .SoundMenuOpen = True
      .SoundItemScroll = True
      .Locked = False
   End With

End Sub

Private Sub Form_Load()

   Call FillMenus
   Call mnbDemo_MenuClick(mnbDemo.StartupMenu)
   
   Show
   DoEvents

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   Call EndDemo

End Sub

Private Sub mnbDemo_CheckMenuWithPassword(MenuIndex As Integer, Cancel As Boolean)

   Call mnbDemo_MenuClick(MenuIndex)
   
   MsgBox "MenuIndex: " & MenuIndex & vbCrLf & vbCrLf & "Is password protected!", vbOKOnly
   Cancel = (InputBox("Enter your password:", "Password") = "")

End Sub

Private Sub mnbDemo_LockMenuWithPassword(MenuIndex As Integer, Cancel As Boolean)

   MsgBox "MenuIndex: " & MenuIndex & vbCrLf & vbCrLf & "Is password protected!" & vbCrLf & vbCrLf & "And will be locked now.", vbOKOnly

End Sub

Private Sub mnbDemo_MenuClick(MenuIndex As Integer)

   lblResult.Item(0).Caption = "MenuIndex: " & MenuIndex

End Sub

Private Sub mnbDemo_MenuItemClick(MenuIndex As Integer, ItemIndex As Integer, ItemKey As String, ItemType As ItemTypes, ItemValue As Boolean, PreviousOption As Integer)

Dim intCount    As Integer
Dim strItemType As String

   strItemType = Choose(ItemType + 1, "DefaultButton", "CheckButton", "OptionButton", "ResetButton", "LockMenuButton")
   
   For intCount = lblResult.Count - 1 To 1 Step -1
      Unload lblResult.Item(intCount)
   Next 'intCount
   
   For intCount = 0 To 5 - (1 And (ItemType <> OptionButton))
      If lblResult.Count < 6 Then
         If intCount Then
            Load lblResult.Item(intCount)
            
            lblResult.Item(intCount).Top = lblResult.Item(intCount - 1).Top + 240
            lblResult.Item(intCount).Visible = True
         End If
      End If
      
      lblResult.Item(intCount).Caption = Choose(intCount + 1, "MenuIndex: ", "  -> Key: ", "  -> ItemIndex: ", "    -> ItemType: ", "        - ItemValue: ", "    -> PreviousOption: ") & Choose(intCount + 1, MenuIndex, ItemKey, ItemIndex, strItemType, ItemValue, PreviousOption)
   Next 'intCount

End Sub

Private Sub mnbMenu_MenuItemClick(MenuIndex As Integer, ItemIndex As Integer, ItemKey As String, ItemType As ItemTypes, ItemValue As Boolean, PreviousOption As Integer)

   If (MenuIndex = 1) And (ItemIndex = 4) Then Call EndDemo

End Sub
