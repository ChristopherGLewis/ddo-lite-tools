VERSION 5.00
Begin VB.UserControl userMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1848
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3096
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1848
   ScaleWidth      =   3096
   Begin VB.Label lblLink 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Command"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   420
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   936
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Caption"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   120
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   684
   End
   Begin VB.Shape shpBorder 
      Height          =   552
      Left            =   1440
      Top             =   120
      Width           =   1032
   End
   Begin VB.Menu mnuMain 
      Caption         =   "List"
      Index           =   0
      Begin VB.Menu mnuList 
         Caption         =   "Edit"
         Index           =   0
      End
      Begin VB.Menu mnuList 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuList 
         Caption         =   "Paste"
         Index           =   2
      End
      Begin VB.Menu mnuList 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuList 
         Caption         =   "Copy List"
         Index           =   4
      End
      Begin VB.Menu mnuList 
         Caption         =   "Delete List"
         Index           =   5
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Command"
      Index           =   1
      Begin VB.Menu mnuCommand 
         Caption         =   "Edit"
         Index           =   0
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "Cut"
         Index           =   2
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "Copy"
         Index           =   3
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "Paste"
         Index           =   4
      End
   End
End
Attribute VB_Name = "userMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event Deleted()
Public Event Copy()
Public Event Changed()

Private Enum CursorEnum
    ceHand = 32649&
    ceSizeAll = 32646&
End Enum

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long

Private mlngForeColor As Long
Private mlngBackColor As Long
Private mlngBorderColor As Long

Private mlngX As Long
Private mlngY As Long
Private mlngMaxLeft As Long
Private mlngMaxTop As Long
Private mlngGridX As Long
Private mlngGridY As Long

Private mblnDragging As Boolean
Private mblnOverride As Boolean

Private mtypMenu As MenuType


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    mlngForeColor = vbWindowText
    mlngBackColor = vbButtonFace
    mlngBorderColor = vbWindowText
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ForeColor", mlngForeColor, vbWindowText
    PropBag.WriteProperty "BackColor", mlngBackColor, vbButtonFace
    PropBag.WriteProperty "BorderColor", mlngBorderColor, vbWindowText
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mlngForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    mlngBackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    mlngBorderColor = PropBag.ReadProperty("BorderColor", vbWindowText)
    DrawControl
End Sub


' ************* USERCONTROL *************


Private Sub UserControl_Initialize()
    Dim lngPixels As Long
    
    With UserControl
        lngPixels = .ScaleY(.TextHeight("Q") \ 2, vbTwips, vbPixels) + 2
        mlngGridY = .ScaleY(lngPixels, vbPixels, vbTwips)
        mlngGridX = .ScaleX(lngPixels, vbPixels, vbTwips)
    End With
    DrawControl
End Sub

Private Sub UserControl_Resize()
    If mblnOverride Then Exit Sub
    DrawControl
End Sub


' ************* PUBLIC *************


Public Sub Init()
    mlngForeColor = cfg.GetColor(cgeWorkspace, cveTextLink)
    mlngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    mlngBorderColor = cfg.GetColor(cgeWorkspace, cveBorderExterior)
    LoadData
End Sub

Public Sub LoadData()
    mtypMenu = gtypMenu
    mtypMenu.LinkList = True
    DrawControl
End Sub

Public Sub SaveData()
    gtypMenu = mtypMenu
End Sub

Public Property Get Deleted() As Boolean
    Deleted = mtypMenu.Deleted
End Property

Public Property Let Deleted(pblnDeleted As Boolean)
    mtypMenu.Deleted = pblnDeleted
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mlngForeColor
End Property

Public Property Let ForeColor(ByVal poleColor As OLE_COLOR)
    mlngForeColor = poleColor
    DrawControl
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mlngBackColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    mlngBackColor = poleColor
    DrawControl
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mlngBorderColor
End Property

Public Property Let BorderColor(ByVal poleColor As OLE_COLOR)
    mlngBorderColor = poleColor
    DrawControl
    PropertyChanged "BorderColor"
End Property


' ************* COMMANDS *************


Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseCursor ceHand
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseCursor ceHand
    mtypMenu.Selected = -1
    Select Case Button
        Case vbLeftButton: EditMenu
        Case vbRightButton: ListMenu True
    End Select
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseCursor ceHand
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    If Y <= UserControl.shpBorder.Top Or mblnDragging Then
        MouseCursor ceSizeAll
        If Button = vbLeftButton And Not mblnOverride Then
            mblnOverride = True
            lngLeft = UserControl.Extender.Left + X - mlngX
            lngTop = UserControl.Extender.Top + Y - mlngY
            lngLeft = (lngLeft \ mlngGridX) * mlngGridX
            lngTop = (lngTop \ mlngGridY) * mlngGridY
            Select Case lngLeft
                Case Is < 0: lngLeft = 0
                Case Is > mlngMaxLeft: lngLeft = mlngMaxLeft
            End Select
            Select Case lngTop
                Case Is < 0: lngTop = 0
                Case Is > mlngMaxTop: lngTop = mlngMaxTop
            End Select
            UserControl.Extender.Move lngLeft, lngTop
            mblnOverride = False
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnMove As Boolean
    
    UserControl.Extender.ZOrder vbBringToFront
    If Y <= UserControl.shpBorder.Top Then
        If Button = vbLeftButton Then MouseCursor ceSizeAll
        blnMove = True
    End If
    Select Case Button
        Case vbLeftButton
            If blnMove Then
                mblnDragging = True
                mblnOverride = False
                With UserControl.Extender.Container
                    mlngMaxLeft = .ScaleWidth - UserControl.Width
                    mlngMaxTop = .ScaleHeight - UserControl.Height
                End With
                mlngX = X
                mlngY = Y
            End If
        Case vbRightButton
            ListMenu False
    End Select
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnDragging Then DirtyFlag dfeLinks
    mblnOverride = False
    mblnDragging = False
    If Y <= UserControl.shpBorder.Top Then MouseCursor ceSizeAll
End Sub

Private Sub lblLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseCursor ceHand
End Sub

Private Sub lblLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseCursor ceHand
    Select Case Button
        Case vbLeftButton
            RunCommand mtypMenu.Command(Index)
        Case vbRightButton
            CommandMenu Index
    End Select
End Sub

Private Sub lblLink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseCursor ceHand
End Sub

Private Sub MouseCursor(penCursor As CursorEnum)
    SetCursor LoadCursor(0, penCursor)
End Sub

Private Sub ListMenu(pblnTitle As Boolean)
    Dim blnVisible As Boolean
    If pblnTitle Then mtypMenu.Selected = -1 Else mtypMenu.Selected = 0
    blnVisible = (gtypClipboard.ClipboardStatus <> cseEmpty)
    UserControl.mnuList(2).Visible = blnVisible
    UserControl.mnuList(3).Visible = blnVisible
    PopupMenu UserControl.mnuMain(0)
End Sub

Private Sub mnuList_Click(Index As Integer)
    Select Case UserControl.mnuList(Index).Caption
        Case "Edit": EditMenu
        Case "Paste": PasteNew
        Case "Copy List": CopyList
        Case "Delete List": DeleteMenu
    End Select
End Sub

Private Sub EditMenu()
    AutoSave
    gtypMenu = mtypMenu
    frmMenuEditor.Show vbModal, frmCompendium
    If gtypMenu.Accepted Then
        mtypMenu = gtypMenu
        RaiseEvent Changed
        DrawControl
    End If
End Sub

Private Sub PasteNew()
    If gtypClipboard.ClipboardStatus = cseEmpty Then Exit Sub
    With mtypMenu
        .Commands = .Commands + 1
        ReDim Preserve .Command(1 To .Commands)
        .Command(.Commands) = gtypClipboard.Command
    End With
    gtypClipboard.ClipboardStatus = cseEmpty
    Change
End Sub

Private Sub DeleteMenu()
    mtypMenu.Deleted = True
    gtypMenu = mtypMenu
    UserControl.Extender.Visible = False
    RaiseEvent Deleted
End Sub

Private Sub CopyList()
    gtypMenu = mtypMenu
    RaiseEvent Copy
End Sub

Private Sub CommandMenu(ByVal plngItem As Long)
    mtypMenu.Selected = plngItem
    UserControl.mnuCommand(4).Enabled = (gtypClipboard.ClipboardStatus <> cseEmpty)
    PopupMenu UserControl.mnuMain(1)
End Sub

Private Sub mnuCommand_Click(Index As Integer)
    Select Case UserControl.mnuCommand(Index).Caption
        Case "Edit": EditMenu
        Case "Cut": CutCommand
        Case "Copy": CopyCommand
        Case "Paste": PasteCommand
    End Select
End Sub

Private Sub CutCommand()
    Dim i As Long
    
    Copy cseCut
    With mtypMenu
        If .Commands = 1 Then
            .Commands = 0
            Erase .Command
        Else
            .Commands = .Commands - 1
            For i = .Selected To .Commands
                .Command(i) = .Command(i + 1)
            Next
            ReDim Preserve .Command(1 To .Commands)
        End If
    End With
    Change
End Sub

Private Sub CopyCommand()
    Copy cseCopied
End Sub

Private Sub Copy(penAction As ClipboardStatusEnum)
    With gtypClipboard
        .ClipboardStatus = penAction
        .Command = mtypMenu.Command(mtypMenu.Selected)
    End With
End Sub

Private Sub PasteCommand()
    Dim i As Long
    
    With mtypMenu
        .Commands = .Commands + 1
        ReDim Preserve .Command(1 To .Commands)
        For i = .Commands To .Selected + 1 Step -1
            .Command(i) = .Command(i - 1)
        Next
        .Command(.Selected) = gtypClipboard.Command
        gtypClipboard.ClipboardStatus = cseEmpty
    End With
    Change
End Sub

Private Sub Change()
    gtypMenu = mtypMenu
    RaiseEvent Changed
    DrawControl
End Sub


' ************* DRAWING *************


Private Sub DrawControl()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngTextHeight As Long
    Dim i As Long
    
    With UserControl
        lngTextHeight = .TextHeight("Q")
        .BackColor = mlngBackColor
        With .lblCaption
            .ForeColor = mlngForeColor
            .BackColor = mlngBackColor
            .Caption = mtypMenu.Title
            '.Width = UserControl.TextWidth(.Caption) + Screen.TwipsPerPixelX * 2
            .Height = lngTextHeight
        End With
        .shpBorder.BorderColor = mlngBorderColor
        With .lblLink(0)
            .ForeColor = mlngForeColor
            .BackColor = mlngBackColor
            .Visible = False
            lngLeft = .Left
            lngTop = .Top
            lngWidth = .Width
            lngHeight = lngTextHeight + Screen.TwipsPerPixelY * 4
        End With
        For i = 1 To mtypMenu.Commands
            If i > .lblLink.UBound Then Load .lblLink(i)
            With .lblLink(i)
                .ForeColor = mlngForeColor
                .BackColor = mlngBackColor
                .Caption = mtypMenu.Command(i).Caption
                .Height = lngTextHeight
                '.Width = UserControl.TextWidth(.Caption) + Screen.TwipsPerPixelX * 2
                If lngWidth < .Width Then lngWidth = .Width
                .Move lngLeft, lngTop
                .Visible = (mtypMenu.Command(i).Style <> mceSeparator)
            End With
            If mtypMenu.Command(i).Style = mceSeparator Then lngTop = lngTop + lngHeight \ 2 Else lngTop = lngTop + lngHeight
        Next
        For i = .lblLink.UBound To i Step -1
            Unload .lblLink(i)
        Next
        mblnOverride = True
        lngWidth = lngWidth + lngLeft * 2
        lngWidth = ((lngWidth + mlngGridX - 1) \ mlngGridX) * mlngGridX
        .Width = lngWidth + Screen.TwipsPerPixelX
        .Height = lngTop + lngHeight
        mblnOverride = False
        SizeBorder
    End With
End Sub

Private Sub SizeBorder()
    Dim lngTop As Long
    
    With UserControl
        lngTop = Screen.TwipsPerPixelY * 10
        .shpBorder.Move 0, lngTop, .ScaleWidth, .ScaleHeight - lngTop
    End With
End Sub
