VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Variables
Private NewMenuItem         As clsMenuItem
Private m_ItemsShowed       As Integer
Private MenuItemsCollection As New Collection

Public Property Get ItemsShowed() As Integer

   ItemsShowed = m_ItemsShowed

End Property

Public Property Let ItemsShowed(ByVal NewItemsShowed As Integer)

   m_ItemsShowed = NewItemsShowed

End Property

Public Function Add(ByVal Caption As String, ByVal Index As Integer, ByVal ButtonHeight As Long, ByRef picIcon As StdPicture) As clsMenuItem

   Set NewMenuItem = New clsMenuItem
   
   With NewMenuItem
      .Caption = Caption
      .Index = Index
      .ButtonHeight = ButtonHeight
      Set .Icon = picIcon
   End With
   
   With MenuItemsCollection
      If (.Count = 0) Or (Index = .Count + 1) Then
         .Add NewMenuItem
         
      ElseIf Index = 1 Then
         .Add NewMenuItem, , 1
         
      Else
         .Add NewMenuItem, , , Index - 1
      End If
   End With
   
   Set Add = NewMenuItem
   Set NewMenuItem = Nothing

End Function

Public Function Count() As Integer

   Count = MenuItemsCollection.Count

End Function

Public Function Item(ByVal Index As Integer) As clsMenuItem

   Set Item = MenuItemsCollection.Item(Index)

End Function

Public Function MouseProcess(ByVal MousePosition As Long, ByVal X As Long, ByVal Y As Long) As Integer

Static intPrevDown As Integer

Dim clsMenuItems   As clsMenuItem

   For Each clsMenuItems In MenuItemsCollection
      With clsMenuItems
         If .HitTest(MousePosition, X, Y) Then
            If MousePosition = MOUSE_UP Then
               If intPrevDown = .Index Then MouseProcess = .Index
               
            Else
               MouseProcess = .Index
            End If
            
            If MousePosition = MOUSE_DOWN Then intPrevDown = .Index
         End If
      End With
   Next 'clsMenuItems
   
   Set clsMenuItems = Nothing

End Function

Public Function Paint(ByVal TopMenuItem As Integer, ByVal FirstIcon As Long, ByVal CurrentMenu As Integer, ByVal ClipHeight As Long, ByVal ItemIconSize As Integer, ByVal OnlyFullItemsShow As Boolean, ByVal OnlyFullItemsHit As Boolean, ByVal ButtonHideInSingleMenu As Boolean) As Boolean

Dim blnIsShowed  As Boolean
Dim clsMenuItems As clsMenuItem

   For Each clsMenuItems In MenuItemsCollection
      With clsMenuItems
         If .Index < TopMenuItem Then
            .DisableButton
            
         Else
            Paint = .PaintButton(TopMenuItem, FirstIcon, CurrentMenu, ClipHeight, ItemIconSize, OnlyFullItemsShow, OnlyFullItemsHit, ButtonHideInSingleMenu, blnIsShowed)
            m_ItemsShowed = m_ItemsShowed + Abs(blnIsShowed)
         End If
      End With
   Next 'clsMenuItems
   
   Set clsMenuItems = Nothing

End Function

Public Sub Delete(ByVal Index As Integer)

   MenuItemsCollection.Remove Index

End Sub

Private Sub Class_Initialize()

   Set MenuItemsCollection = New Collection

End Sub

Private Sub Class_Terminate()

   Set NewMenuItem = Nothing
   Set MenuItemsCollection = Nothing

End Sub
