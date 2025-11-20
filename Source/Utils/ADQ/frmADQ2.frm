VERSION 5.00
Begin VB.Form frmADQ2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Against the Demon Queen"
   ClientHeight    =   552
   ClientLeft      =   3240
   ClientTop       =   0
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0039A3C5&
   Icon            =   "frmADQ2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   552
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optClose 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5190
      Picture         =   "frmADQ2.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
      Width           =   315
   End
   Begin VB.OptionButton optCopy 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4860
      Picture         =   "frmADQ2.frx":058E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Copy to Clipboard"
      Top             =   15
      Width           =   315
   End
   Begin VB.Timer tmrAlwaysOnTop 
      Interval        =   1000
      Left            =   180
      Top             =   120
   End
End
Attribute VB_Name = "frmADQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

' RGB = 197,163,57

Private mblnMoving As Boolean
Private msngX As Single
Private msngY As Single
Private msngMaxX As Single
Private msngMaxY As Single

Private mstrSolution As String
Private mstrLine() As String


' ************* FORM *************


Private Sub Form_Load()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    GetDesktop lngLeft, lngTop, lngWidth, lngHeight
    msngMaxX = lngWidth - Me.Width
    msngMaxY = lngHeight - Me.Height
    
'    Me.lblADQ.ForeColor = lngColor
'    Me.shpBorder.BorderColor = lngColor
    PositionForm Me, "ADQ2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteCoords "ADQ2", Me.Left, Me.Top
    CloseApp
End Sub

Private Sub tmrAlwaysOnTop_Timer()
    SetAlwaysOnTop Me.hWnd, True
End Sub

Public Property Let Solution(pstrSolution As String)
    mstrLine = Split(pstrSolution, vbNewLine)
    mstrSolution = Join(mstrLine, "  ")
    DrawSolution
    Clipboard.Clear
    Clipboard.SetText mstrSolution
End Property

Private Sub DrawSolution()
    Dim lngColor As Long
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    lngColor = RGB(197, 163, 57)
    Me.Cls
    lngLeft = Me.ScaleX(12, vbPixels, vbTwips)
    Me.CurrentX = lngLeft
    Me.CurrentY = (Me.ScaleHeight - Me.TextHeight("Q") * 2) \ 2 - Screen.TwipsPerPixelY
    Me.Print mstrLine(0)
    Me.CurrentX = lngLeft
    Me.Print mstrLine(1)
    ' Size form
    lngWidth = Me.TextWidth(mstrLine(0))
    If lngWidth < Me.TextWidth(mstrLine(1)) Then lngWidth = Me.TextWidth(mstrLine(1))
    Me.Width = lngWidth + lngLeft * 2 + Me.optCopy.Width + Me.optClose.Width
    Me.optClose.Left = Me.ScaleWidth - Me.optClose.Width - Screen.TwipsPerPixelX
    Me.optCopy.Left = Me.optClose.Left - Me.optClose.Width - Screen.TwipsPerPixelX
    ' Draw border
    Me.Line (0, 0)-(Me.ScaleWidth - Screen.TwipsPerPixelX, Me.ScaleHeight - Screen.TwipsPerPixelY), lngColor, B
End Sub


' ************* MOVE *************


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngX = X
    msngY = Y
    mblnMoving = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngX As Single
    Dim sngY As Single
    
    If Not mblnMoving Then Exit Sub
    
    sngX = Me.Left - (msngX - X) \ Screen.TwipsPerPixelX
    Select Case sngX
        Case Is < 0: sngX = 0
        Case Is > msngMaxX: sngX = msngMaxX
    End Select
    
    sngY = Me.Top - (msngY - Y) \ Screen.TwipsPerPixelY
    Select Case sngY
        Case Is < 0: sngY = 0
        Case Is > msngMaxY: sngY = msngMaxY
    End Select
    
    Me.Move sngX, sngY
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMoving = False
End Sub


' ************* BUTTONS *************


Private Sub optCopy_Click()
    Me.optCopy.Value = False
    Clipboard.Clear
    Clipboard.SetText mstrSolution
End Sub

Private Sub optClose_Click()
    Me.optCopy.Value = False
    Unload Me
End Sub
